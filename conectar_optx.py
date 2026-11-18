"""Localiza o cenário HO mais recente, solicita a exportação no OPT+Z e baixa o XLSX."""

from __future__ import annotations
from excel_io import carregar_cenario
import io
import time
import unicodedata
import zipfile
from pathlib import Path
from typing import Any

import requests


# ============================================================
# CONFIGURAÇÕES
# ============================================================

URL_LOGIN = "http://optz2.integra.itg:5002/signin"
URL_API = "http://optz2.integra.itg:5000"
URL_EXPORTACOES = "http://optz2.integra.itg:5005"

CONTA = "jca"
LIMIT_PLANEJAMENTOS = 10
LIMIT_CENARIOS = 20
LIMIT_EXPORTACOES = 50
TIMEOUT = 30
TIMEOUT_DOWNLOAD = 120
INTERVALO_CONSULTA = 5
TEMPO_MAXIMO_ESPERA = 1500  # 25 minutos
MODELO_EXPORTACAO = 1

PASTA_EXPORTACAO = Path.cwd() / "exportacoes_optz"


# ============================================================
# EXCEÇÕES E UTILITÁRIOS
# ============================================================

class ExportacaoFalhou(RuntimeError):
    """A plataforma informou que a exportação falhou ou foi cancelada."""


class ExportacaoNaoPronta(RuntimeError):
    """A API ainda não disponibilizou o arquivo para download."""


def normalizar_texto(valor: Any) -> str:
    texto = str(valor or "").strip().casefold()
    return "".join(
        caractere
        for caractere in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(caractere)
    )


def obter_json(resposta: requests.Response, contexto: str) -> dict:
    try:
        dados = resposta.json()
    except ValueError as erro:
        trecho = resposta.text[:500] if resposta.text else "<resposta vazia>"
        raise ValueError(
            f"{contexto}: a API não retornou JSON válido. Resposta: {trecho}"
        ) from erro

    if not isinstance(dados, dict):
        raise ValueError(f"{contexto}: formato JSON inesperado.")
    return dados


def extrair_id_exportacao(dados: dict) -> str | None:
    """Procura o identificador da exportação sem confundi-lo com contaId."""
    if not isinstance(dados, dict):
        return None

    for chave in ("id", "exportacaoId", "exportacao_id"):
        valor = dados.get(chave)
        if isinstance(valor, str) and valor.strip():
            return valor.strip()

    # Algumas respostas podem envolver o objeto em outra propriedade.
    for chave in ("data", "exportacao", "registro"):
        interno = dados.get(chave)
        if isinstance(interno, dict):
            encontrado = extrair_id_exportacao(interno)
            if encontrado:
                return encontrado

    return None


def nome_arquivo_seguro(valor: str) -> str:
    nome = "".join(
        caractere if caractere.isalnum() or caractere in "-_" else "_"
        for caractere in str(valor or "CENARIO_HO")
    )
    return nome or "CENARIO_HO"


# ============================================================
# AUTENTICAÇÃO
# ============================================================

def autenticar(session: requests.Session) -> None:
    print("\n========== AUTENTICAÇÃO OPT+Z ==========")
    usuario = input("Usuário OPT+Z: ").strip()
    senha = input("Senha OPT+Z: ")

    if not usuario or not senha:
        raise ValueError("Usuário e senha devem ser informados.")

    resposta = session.post(
        URL_LOGIN,
        json={"conta": CONTA, "usuario": usuario, "senha": senha},
        headers={"Accept": "application/json"},
        timeout=TIMEOUT,
    )
    resposta.raise_for_status()
    retorno = obter_json(resposta, "Autenticação")

    dados = retorno.get("data") or {}
    if not isinstance(dados, dict):
        raise ValueError("Formato inesperado no retorno da autenticação.")

    access_token = dados.get("accessToken")
    if not access_token:
        raise ValueError("A API não retornou accessToken.")

    session.headers.update({
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    })
    print("Autenticação realizada com sucesso.")


# ============================================================
# CONSULTAR PLANEJAMENTOS
# ============================================================

def buscar_planejamentos(session: requests.Session) -> list[dict]:
    print("\n========== CONSULTANDO PLANEJAMENTOS ==========")
    planejamentos: list[dict] = []
    pagina = 1
    total = None

    while True:
        parametros = {
            "limit": LIMIT_PLANEJAMENTOS,
            "page": pagina,
            "pesquisa": "",
            "sort": "nome:asc",
        }
        resposta = session.get(
            f"{URL_API}/planejamentos", params=parametros, timeout=TIMEOUT
        )
        resposta.raise_for_status()
        retorno = obter_json(resposta, "Consulta de planejamentos")

        registros = retorno.get("data", [])
        total = retorno.get("total", total)
        if not isinstance(registros, list):
            raise ValueError("A lista de planejamentos não veio em formato de lista.")
        if not registros:
            break

        planejamentos.extend(registros)
        print(f"Página {pagina}: {len(registros)} planejamentos recebidos.")

        if total is not None and len(planejamentos) >= int(total):
            break
        if len(registros) < LIMIT_PLANEJAMENTOS:
            break
        pagina += 1

    print(f"Total de planejamentos carregados: {len(planejamentos)}")
    return planejamentos


def selecionar_planejamento(planejamentos: list[dict]) -> dict:
    if not planejamentos:
        raise ValueError("Nenhum planejamento encontrado.")

    print("\n========== PLANEJAMENTOS DISPONÍVEIS ==========")
    for indice, planejamento in enumerate(planejamentos, start=1):
        empresa = (planejamento.get("empresa") or {}).get("codigo") or ""
        print(
            f"{indice:02d} - {planejamento.get('nome', 'Sem nome')} | "
            f"Empresa: {empresa} | "
            f"Cenários: {planejamento.get('numCenarios', 'Não informado')}"
        )
        descricao = planejamento.get("descricao")
        if descricao:
            print(f"     {descricao}")

    while True:
        try:
            escolha = int(input("\nSelecione o número do planejamento: ").strip())
        except ValueError:
            print("Digite um número inteiro válido.")
            continue

        if 1 <= escolha <= len(planejamentos):
            selecionado = planejamentos[escolha - 1]
            break
        print(f"Escolha um número entre 1 e {len(planejamentos)}.")

    if not selecionado.get("id"):
        raise ValueError("O planejamento selecionado não possui ID.")

    print("\nPlanejamento selecionado:", selecionado.get("nome"))
    print("ID:", selecionado["id"])
    return selecionado


# ============================================================
# LOCALIZAR O CENÁRIO HO MAIS RECENTE
# ============================================================

def buscar_ultimo_cenario_ho(
    session: requests.Session, planejamento_id: str
) -> dict:
    print("\n========== LOCALIZANDO O CENÁRIO HO MAIS RECENTE ==========")
    pagina = 1

    while True:
        parametros = {
            "limit": LIMIT_CENARIOS,
            "page": pagina,
            "pesquisa": "",
            "planejamentoId": planejamento_id,
            "sort": "criadoEm:desc",
        }
        resposta = session.get(
            f"{URL_API}/cenarios", params=parametros, timeout=TIMEOUT
        )
        resposta.raise_for_status()
        retorno = obter_json(resposta, "Consulta de cenários")

        cenarios = retorno.get("data", [])
        total = retorno.get("total")
        if not isinstance(cenarios, list):
            raise ValueError("A lista de cenários não veio em formato de lista.")
        if not cenarios:
            break

        print(f"Verificando página {pagina}: {len(cenarios)} cenários.")
        for cenario in cenarios:
            codigo = str(cenario.get("codigo") or "").strip()
            if codigo.upper().endswith("HO"):
                if not cenario.get("id"):
                    raise ValueError("O cenário HO encontrado não possui ID.")
                print("\nCenário HO encontrado:")
                print("Código:", codigo)
                print("Descrição:", cenario.get("descricao"))
                print("Criado em:", cenario.get("criadoEm"))
                print("ID:", cenario["id"])
                return cenario

        if total is not None and pagina * LIMIT_CENARIOS >= int(total):
            break
        if len(cenarios) < LIMIT_CENARIOS:
            break
        pagina += 1

    raise ValueError(
        "Nenhum cenário com código terminado em HO foi encontrado "
        "no planejamento selecionado."
    )


# ============================================================
# SOLICITAR EXPORTAÇÃO OFICIAL
# ============================================================

def solicitar_exportacao(
    session: requests.Session, cenario_id: str
) -> tuple[str, dict]:
    print("\n========== SOLICITANDO EXPORTAÇÃO OFICIAL ==========")
    url = f"{URL_EXPORTACOES}/exportacoes/entidade/cenario"
    parametros = {"confirmed": "false"}
    payload = {"cenarioId": cenario_id, "modelo": MODELO_EXPORTACAO}

    resposta = session.post(
        url, params=parametros, json=payload, timeout=TIMEOUT
    )
    resposta.raise_for_status()
    retorno = obter_json(resposta, "Solicitação de exportação")

    dados = retorno.get("data") or {}
    if not isinstance(dados, dict):
        raise ValueError("Formato inesperado nos dados da exportação.")

    exportacao_id = extrair_id_exportacao(dados)
    print("Solicitação enviada.")
    print("Status HTTP:", resposta.status_code)
    print("Status informado pela API:", retorno.get("statusCode"))
    print("Campos retornados:", list(dados.keys()))

    if not exportacao_id:
        raise ValueError(
            "A resposta do POST não apresentou um ID de exportação reconhecível. "
            "Confira os campos retornados acima e a listagem GET /exportacoes."
        )

    print("ID da exportação:", exportacao_id)
    return exportacao_id, dados


# ============================================================
# CONSULTAR STATUS NA LISTAGEM DE EXPORTAÇÕES
# ============================================================

def buscar_exportacao_na_lista(
    session: requests.Session, exportacao_id: str
) -> dict | None:
    """Busca o registro na rota GET /exportacoes (não baixa o binário)."""
    pagina = 1

    while True:
        parametros = {
            "limit": LIMIT_EXPORTACOES,
            "page": pagina,
            "pesquisa": "",
            "sort": "criadoEm:desc",
        }
        resposta = session.get(
            f"{URL_EXPORTACOES}/exportacoes",
            params=parametros,
            timeout=TIMEOUT,
        )
        resposta.raise_for_status()
        retorno = obter_json(resposta, "Consulta da lista de exportações")

        registros = retorno.get("data", [])
        total = retorno.get("total")
        if not isinstance(registros, list):
            raise ValueError("A lista de exportações não veio em formato de lista.")

        for registro in registros:
            if not isinstance(registro, dict):
                continue
            registro_id = extrair_id_exportacao(registro)
            if registro_id == exportacao_id:
                return registro

        if not registros or len(registros) < LIMIT_EXPORTACOES:
            return None
        if total is not None and pagina * LIMIT_EXPORTACOES >= int(total):
            return None
        pagina += 1


def avaliar_status_exportacao(registro: dict) -> tuple[str, str]:
    """Retorna estado apenas informativo; status numérico não autoriza download."""
    status = registro.get("status")
    progresso = registro.get("progress", registro.get("progresso"))
    resumo = f"status={status!r}, progress={progresso!r}"

    # Não interpretar status=2 como concluído: nesta instalação, a listagem
    # já apresentou status=2/progress=0 enquanto o endpoint de arquivo ainda
    # respondeu HTTP 400 com 'Exportação não concluída'.
    status_texto = normalizar_texto(status) if isinstance(status, str) else ""
    falhas = ("erro", "error", "falhou", "failed", "cancelado", "cancelada", "cancelled")
    concluidos = (
        "concluida", "concluido", "completed", "finished", "finalizada",
        "finalizado", "done", "success", "succeeded",
    )

    if any(palavra in status_texto for palavra in falhas):
        return "falhou", resumo
    if any(palavra in status_texto for palavra in concluidos):
        return "indicada_concluida", resumo

    return "nao_confirmada", resumo


def extrair_mensagem_erro(resposta: requests.Response) -> str:
    """Extrai mensagens de erro JSON ou texto sem presumir o formato."""
    try:
        dados = resposta.json()
    except ValueError:
        return resposta.text[:500] if resposta.text else ""

    if isinstance(dados, dict):
        erros = dados.get("errors")
        if isinstance(erros, list):
            return "; ".join(str(item) for item in erros)
        if erros:
            return str(erros)
        return str(dados.get("message") or dados.get("mensagem") or dados)
    return str(dados)


def aguardar_exportacao_concluida(
    session: requests.Session, exportacao_id: str, codigo_cenario: str
) -> Path:
    """Só considera concluída quando o endpoint entrega um XLSX válido.

    A listagem GET /exportacoes é consultada apenas para diagnóstico. O código
    numérico de status não é usado como confirmação final, pois o endpoint de
    download pode continuar retornando 'Exportação não concluída'.
    """
    print("\n========== AGUARDANDO ARQUIVO DA EXPORTAÇÃO ==========")
    inicio = time.monotonic()
    ultimo_estado = "status ainda não consultado"

    while True:
        decorrido = int(time.monotonic() - inicio)

        # Consultar a listagem apenas para exibir o estado reportado pelo site.
        try:
            registro = buscar_exportacao_na_lista(session, exportacao_id)
            if registro:
                estado, resumo = avaliar_status_exportacao(registro)
                ultimo_estado = resumo
                print(f"Exportação na listagem: {resumo}")
                if estado == "falhou":
                    raise ExportacaoFalhou(
                        f"A plataforma informou falha/cancelamento: {resumo}"
                    )
            else:
                print("Exportação ainda não localizada na listagem.")
        except ExportacaoFalhou:
            raise
        except (requests.RequestException, ValueError) as erro:
            # A listagem é auxiliar. O endpoint de download será a confirmação.
            print(f"Não foi possível atualizar a listagem: {erro}")

        # Tentativa real de download. Só retorna quando receber XLSX válido.
        try:
            return baixar_exportacao(
                session=session,
                exportacao_id=exportacao_id,
                codigo_cenario=codigo_cenario,
            )
        except ExportacaoNaoPronta as erro:
            print(
                f"Arquivo ainda indisponível ({decorrido}s): {erro}. "
                "Nova tentativa em "
                f"{INTERVALO_CONSULTA}s..."
            )

        if decorrido >= TEMPO_MAXIMO_ESPERA:
            raise TimeoutError(
                f"O OPT+Z ainda não disponibilizou o arquivo após "
                f"{TEMPO_MAXIMO_ESPERA} segundos. "
                f"Último estado da listagem: {ultimo_estado}. "
                "O script não salvou nenhum arquivo incompleto."
            )

        time.sleep(INTERVALO_CONSULTA)


# ============================================================
# BAIXAR O ARQUIVO XLSX
# ============================================================

def baixar_exportacao(
    session: requests.Session, exportacao_id: str, codigo_cenario: str
) -> Path:
    """Tenta obter o XLSX; sinaliza quando o backend ainda não o liberou."""
    url = f"{URL_EXPORTACOES}/exportacoes/{exportacao_id}"
    headers = {
        "Accept": (
            "application/octet-stream, "
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, */*"
        )
    }

    resposta = session.get(
        url, headers=headers, timeout=TIMEOUT_DOWNLOAD
    )

    # O próprio backend informa explicitamente quando o arquivo ainda não está pronto.
    if resposta.status_code in (202, 404, 409, 425):
        mensagem = extrair_mensagem_erro(resposta)
        raise ExportacaoNaoPronta(
            f"HTTP {resposta.status_code}: {mensagem or 'arquivo indisponível'}"
        )

    if resposta.status_code == 400:
        mensagem = extrair_mensagem_erro(resposta)
        mensagem_normalizada = normalizar_texto(mensagem)
        if "exportacao nao concluida" in mensagem_normalizada or \
           "arquivo nao disponivel" in mensagem_normalizada or \
           "ainda processando" in mensagem_normalizada:
            raise ExportacaoNaoPronta(
                f"HTTP 400: {mensagem or 'exportação não concluída'}"
            )
        resposta.raise_for_status()

    resposta.raise_for_status()
    conteudo = resposta.content

    if not conteudo:
        raise ExportacaoNaoPronta("a resposta do download veio vazia")

    # Validar o pacote XLSX antes de gravar no disco.
    try:
        with zipfile.ZipFile(io.BytesIO(conteudo)) as arquivo_zip:
            nomes = set(arquivo_zip.namelist())
            if "xl/workbook.xml" not in nomes:
                raise ValueError(
                    "O servidor retornou um ZIP, mas não um Excel XLSX. "
                    f"Primeiros arquivos: {list(nomes)[:10]}"
                )
            arquivo_corrompido = arquivo_zip.testzip()
            if arquivo_corrompido:
                raise ValueError(
                    f"O arquivo XLSX parece corrompido: {arquivo_corrompido}"
                )
    except zipfile.BadZipFile as erro:
        trecho = conteudo[:300].decode("utf-8", errors="replace")
        mensagem_normalizada = normalizar_texto(trecho)
        if "exportacao nao concluida" in mensagem_normalizada:
            raise ExportacaoNaoPronta(trecho) from erro
        raise ValueError(
            "A rota GET /exportacoes/{id} não retornou um XLSX válido. "
            f"Content-Type: {resposta.headers.get('Content-Type')!r}. "
            f"Início do conteúdo: {trecho!r}"
        ) from erro

    PASTA_EXPORTACAO.mkdir(parents=True, exist_ok=True)
    codigo_seguro = nome_arquivo_seguro(codigo_cenario)
    caminho = PASTA_EXPORTACAO / f"CENARIO_{codigo_seguro}.xlsx"
    caminho.write_bytes(conteudo)

    print("\n========== DOWNLOAD CONCLUÍDO ==========")
    print("Content-Type:", resposta.headers.get("Content-Type"))
    print("Content-Disposition:", resposta.headers.get("Content-Disposition"))
    print("Tamanho:", len(conteudo), "bytes")
    print("Arquivo salvo em:", caminho.resolve())
    return caminho


# ============================================================
# EXECUÇÃO PRINCIPAL
# ============================================================

def main() -> None:
    with requests.Session() as session:
        autenticar(session)

        planejamentos = buscar_planejamentos(session)
        planejamento = selecionar_planejamento(planejamentos)

        cenario = buscar_ultimo_cenario_ho(
            session=session,
            planejamento_id=planejamento["id"],
        )

        # Solicita ao próprio OPT+Z a geração do arquivo oficial.
        exportacao_id, _dados_exportacao = solicitar_exportacao(
            session=session,
            cenario_id=cenario["id"],
        )

        # A listagem mostra o status, mas não é considerada prova de que
        # o arquivo esteja liberado. O download é tentado repetidamente até
        # o endpoint devolver um XLSX válido.
        caminho = aguardar_exportacao_concluida(
            session=session,
            exportacao_id=exportacao_id,
            codigo_cenario=cenario.get("codigo") or "HO",
        )

        # ========================================================
        # IMPORTAR CENÁRIO EXPORTADO PARA O PYTHON
        # ========================================================

        print("\n========== IMPORTANDO CENÁRIO PARA O PYTHON ==========")

        df_cenario, bgs_cenario, fontes_cenario = carregar_cenario(
            str(caminho)
        )

        print("Cenário importado com sucesso!")
        print("Quantidade de linhas:", len(df_cenario))
        print("Colunas:", list(df_cenario.columns))

        print("\n========== PROCESSO FINALIZADO ==========")
        print("Planejamento:", planejamento.get("nome"))
        print("Cenário:", cenario.get("codigo"))
        print("Arquivo:", caminho.resolve())
        print("Nenhum dado do cenário foi alterado no OPT+Z.")


if __name__ == "__main__":
    try:
        main()
    except requests.exceptions.HTTPError as erro:
        print("\nErro HTTP:", erro)
        if erro.response is not None:
            print("Resposta:", erro.response.text[:1000])
    except requests.exceptions.Timeout:
        print("\nTempo limite da requisição excedido.")
    except requests.exceptions.ConnectionError as erro:
        print("\nErro de conexão:", erro)
    except requests.exceptions.RequestException as erro:
        print("\nErro na requisição:", erro)
    except (ValueError, KeyError, OSError, TimeoutError, ExportacaoFalhou) as erro:
        print("\nErro:", erro)