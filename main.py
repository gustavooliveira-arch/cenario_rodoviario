import sys
import time
import traceback
import copy
import unicodedata
import io
import zipfile
import requests
from collections import Counter
from flask import Flask
from estilo_app import (
    aplicar_estilo,
    aplicar_fundo_blur,
    DialogoProgressoLinha,
    QInputDialog,
    QMessageBox,
)

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

from PySide6.QtCore import (
    QAbstractTableModel,
    QModelIndex,
    QObject,
    QPoint,
    QPointF,
    QRect,
    QRectF,
    QSortFilterProxyModel,
    Qt,
    QThread,
    QTimer,
    Signal,
    Slot,
)

from PySide6.QtGui import (
    QBrush,
    QColor,
    QFontMetrics,
    QIcon,
    QPainter,
    QPen,
    QPolygonF,
)

from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QStyle,
    QStyledItemDelegate,
    QTableView,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from excel_io import (
    carregar_cenario,
    carregar_timetable,
    salvar_cenario,
    cor_preenchimento,
    cor_fonte,
    paleta_tema,
)

from regras_cenario import _processar as processar_em_memoria


PASTA_APP = Path(__file__).resolve().parent
CAMINHO_ICONE = PASTA_APP / "assets" / "logo.ico"


# ============================================================
# RESULTADO DO PROCESSAMENTO
# ============================================================

@dataclass
class ResultadoProcessamento:
    # Dados exibidos no aplicativo, sem TIPO DE ALTERAÇÃO
    df: object
    bgs: list
    fontes: list

    # Dados completos para exportação Excel
    # TIPO DE ALTERAÇÃO será a primeira coluna
    df_export: object
    bgs_export: list
    fontes_export: list

    relatorio: str


# ============================================================
# CONVERSÃO DE VALORES
# ============================================================

def _valor_excel(valor):
    """Converte valores do pandas para tipos aceitos pelo openpyxl."""

    if valor is None or isinstance(valor, str):
        return valor

    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass

    # Tipos numpy/pandas escalares são convertidos para tipos Python.
    if hasattr(valor, "item") and callable(valor.item):
        try:
            valor = valor.item()
        except (ValueError, TypeError):
            pass

    return valor


def _cor_hex(valor, padrao="#FFFFFF"):
    if valor is None:
        return padrao

    if hasattr(valor, "name") and callable(valor.name):
        try:
            valor = valor.name()
        except Exception:
            pass

    texto = str(valor).strip()

    if texto.startswith("#") and len(texto) == 7:
        return texto.upper()

    if len(texto) == 8 and texto.startswith("FF"):
        return "#" + texto[-6:].upper()

    return padrao


def _normalizar_cor(valor, padrao=None):
    """Converte qualquer cor lida do Excel para #RRGGBB.

    Usa a mesma regra das células (últimos 6 dígitos do ARGB/RGB), para que
    cabeçalhos e células mostrem exatamente a mesma cor do arquivo.
    """
    if valor is None:
        return padrao

    if hasattr(valor, "name") and callable(valor.name):
        try:
            valor = valor.name()
        except Exception:
            pass

    texto = str(valor).strip().lstrip("#")

    if len(texto) in (6, 8) and all(
        c in "0123456789abcdefABCDEF" for c in texto
    ):
        return "#" + texto[-6:].upper()

    return padrao


def _normalizar_texto(valor):
    """Texto para comparação de filtros: sem acento e sem diferença de caixa."""
    if valor is None:
        return ""

    try:
        if valor != valor:  # NaN
            return ""
    except Exception:
        pass

    texto = str(valor)

    if not texto.isascii():
        texto = unicodedata.normalize("NFKD", texto)
        texto = "".join(c for c in texto if not unicodedata.combining(c))

    return texto.casefold()


def _ler_estilos_cabecalho_excel(caminho, nome_aba, linha_cabecalho, quantidade_colunas):
    """Lê as cores reais do cabeçalho no Excel, incluindo cores de tema/tint.

    As cores das células já vêm de carregar_cenario/carregar_timetable
    (excel_io), que usam as mesmas funções cor_preenchimento e cor_fonte.
    """
    if not caminho or not nome_aba or not linha_cabecalho or quantidade_colunas <= 0:
        return [], []

    wb = load_workbook(caminho, data_only=True)
    try:
        if nome_aba not in wb.sheetnames:
            return [], []

        ws = wb[nome_aba]
        tema = paleta_tema(wb)
        fundos = []
        fontes = []

        for coluna in range(1, quantidade_colunas + 1):
            celula = ws.cell(row=int(linha_cabecalho), column=coluna)
            fundos.append(cor_preenchimento(celula, tema) or "#FFFFFF")
            # Quando a fonte não tem cor explícita, o padrão do Excel é preto.
            fontes.append(cor_fonte(celula, tema) or "#000000")

        return fundos, fontes
    finally:
        wb.close()


def _aplicar_estilos_cabecalho(dados, caminho):
    """Acrescenta a dados as cores do cabeçalho lidas do Excel."""
    dados["header_bgs"], dados["header_fontes"] = _ler_estilos_cabecalho_excel(
        caminho,
        dados["aba"],
        dados["header_row"],
        len(dados["df"].columns),
    )


# ============================================================
# NORMALIZAÇÃO E SEPARAÇÃO DA COLUNA TIPO DE ALTERAÇÃO
# ============================================================

def _normalizar_nome_coluna(valor):
    """
    Normaliza o nome da coluna para comparação,
    ignorando acentos e diferenças de espaços.
    """

    texto = unicodedata.normalize(
        "NFKD",
        str(valor)
    )

    texto = "".join(
        c for c in texto
        if not unicodedata.combining(c)
    )

    return " ".join(texto.upper().split())


def _separar_coluna_tipo_alteracao(df, bgs, fontes):
    """
    Cria duas versões do resultado:

    1. Aplicativo:
       Sem a coluna TIPO DE ALTERAÇÃO.

    2. Excel:
       Com TIPO DE ALTERAÇÃO como primeira coluna.

    Preserva os valores, cores e fontes correspondentes.
    """

    indice = next(
        (
            i
            for i, nome in enumerate(df.columns)
            if _normalizar_nome_coluna(nome) == "TIPO DE ALTERACAO"
        ),
        None
    )

    # Se a coluna não existir, mantém o resultado original.
    if indice is None:

        df_visual = df.copy()
        df_export = df.copy()

        bgs_visual = [
            list(linha)
            for linha in bgs
        ]

        fontes_visual = [
            list(linha)
            for linha in fontes
        ]

        bgs_export = [
            list(linha)
            for linha in bgs
        ]

        fontes_export = [
            list(linha)
            for linha in fontes
        ]

        return (
            df_visual,
            bgs_visual,
            fontes_visual,
            df_export,
            bgs_export,
            fontes_export
        )

    # --------------------------------------------------------
    # VERSÃO DO APLICATIVO: remove TIPO DE ALTERAÇÃO
    # --------------------------------------------------------

    colunas_visiveis = [
        i
        for i in range(df.shape[1])
        if i != indice
    ]

    df_visual = df.iloc[:, colunas_visiveis].copy()

    bgs_visual = [
        [
            linha[i] if i < len(linha) else "#FFFFFF"
            for i in colunas_visiveis
        ]
        for linha in bgs
    ]

    fontes_visual = [
        [
            linha[i] if i < len(linha) else None
            for i in colunas_visiveis
        ]
        for linha in fontes
    ]

    # --------------------------------------------------------
    # VERSÃO DO EXCEL: TIPO DE ALTERAÇÃO na primeira coluna
    # --------------------------------------------------------

    colunas_exportacao = [
        indice
    ] + colunas_visiveis

    df_export = df.iloc[:, colunas_exportacao].copy()

    bgs_export = [
        [
            linha[i] if i < len(linha) else "#FFFFFF"
            for i in colunas_exportacao
        ]
        for linha in bgs
    ]

    fontes_export = [
        [
            linha[i] if i < len(linha) else None
            for i in colunas_exportacao
        ]
        for linha in fontes
    ]

    return (
        df_visual,
        bgs_visual,
        fontes_visual,
        df_export,
        bgs_export,
        fontes_export
    )


# ============================================================
# MONTAGEM DAS ABAS EM MEMÓRIA
# ============================================================

def _montar_aba(ws, df, bgs=None, fontes=None):
    """Escreve DataFrame e cores em uma aba com cabeçalho na primeira linha."""

    if df is None:
        raise ValueError(
            "Uma das planilhas não foi importada corretamente."
        )

    for c, nome in enumerate(df.columns, 1):
        ws.cell(
            1,
            c,
            "" if nome is None else str(nome)
        )

    valores = df.astype(object).values.tolist()

    cache_fill = {}
    cache_fonte = {}

    for r, linha in enumerate(valores):

        bg_linha = (
            bgs[r]
            if bgs and r < len(bgs)
            else None
        )

        ft_linha = (
            fontes[r]
            if fontes and r < len(fontes)
            else None
        )

        for c, valor in enumerate(linha):

            cell = ws.cell(
                r + 2,
                c + 1
            )

            valor = _valor_excel(valor)

            if valor is not None:
                cell.value = valor

            # Branco = sem preenchimento.
            if bg_linha is not None and c < len(bg_linha):

                cor = _cor_hex(bg_linha[c])

                if cor != "#FFFFFF":

                    fill = cache_fill.get(cor)

                    if fill is None:

                        fill = cache_fill[cor] = PatternFill(
                            fill_type="solid",
                            fgColor="FF" + cor[1:]
                        )

                    cell.fill = fill

            if (
                ft_linha is not None
                and c < len(ft_linha)
                and ft_linha[c]
            ):

                cor = _cor_hex(
                    ft_linha[c],
                    padrao="#000000"
                )

                fonte = cache_fonte.get(cor)

                if fonte is None:

                    fonte = copy.copy(cell.font)

                    fonte.color = "FF" + cor[1:]

                    cache_fonte[cor] = fonte

                cell.font = fonte


# ============================================================
# LEITURA DA ABA PROCESSADA
# ============================================================

def _ler_aba_processada(ws):
    """Lê valores, preenchimentos e cores de fonte do resultado das regras."""

    wb = ws.parent

    max_row = ws.max_row
    max_col = ws.max_column

    # Cor de cada preenchimento, calculada uma vez por estilo.
    fill_hex = []

    for f in wb._fills:

        cor = "#FFFFFF"

        fg = getattr(f, "fgColor", None)

        if (
            getattr(f, "fill_type", None) == "solid"
            and fg is not None
            and fg.type == "rgb"
        ):
            cor = "#" + str(fg.rgb)[-6:]

        fill_hex.append(cor)

    # Cor de cada fonte.
    font_hex = []

    for f in wb._fonts:

        c = f.color

        font_hex.append(
            "#"
            + str(c.rgb)[-6:]
            if c is not None and c.type == "rgb"
            else None
        )

    # Acesso direto às células existentes.
    cells = ws._cells

    intervalo = range(
        1,
        max_col + 1
    )

    cabecalho = []

    for c in intervalo:

        cell = cells.get((1, c))

        cabecalho.append(
            cell.value
            if cell is not None
            else None
        )

    linhas = []
    bgs = []
    fontes = []

    for r in range(2, max_row + 1):

        valores = []
        cores_linha = []
        fontes_linha = []

        vazia = True

        for c in intervalo:

            cell = cells.get((r, c))

            if cell is None:

                valores.append(None)
                cores_linha.append("#FFFFFF")
                fontes_linha.append(None)

                continue

            v = cell.value

            if v is not None:
                vazia = False

            st = cell._style

            valores.append(v)

            if st is None:

                # Célula sem formatação explícita.
                cores_linha.append("#FFFFFF")
                fontes_linha.append(None)

            else:

                fill_id = st.fillId
                font_id = st.fontId

                cores_linha.append(
                    fill_hex[fill_id]
                    if fill_id < len(fill_hex)
                    else "#FFFFFF"
                )

                fontes_linha.append(
                    font_hex[font_id]
                    if font_id < len(font_hex)
                    else None
                )

        if vazia:
            continue

        linhas.append(valores)
        bgs.append(cores_linha)
        fontes.append(fontes_linha)

    return (
        pd.DataFrame(
            linhas,
            columns=cabecalho
        ),
        bgs,
        fontes
    )


# ============================================================
# PROCESSAMENTO DO CENÁRIO
# ============================================================

def processar_cenario(
    df_tt,
    bg_tt,
    df_cen,
    bg_cen,
    cols_tt=None,
    cols_cen=None,
    fontes_cen=None,
    progresso=None
):
    """
    Monta as duas abas em um workbook em memória e aplica as regras.

    Mantém duas versões do resultado:

    - Versão visual, sem TIPO DE ALTERAÇÃO.
    - Versão de exportação, com TIPO DE ALTERAÇÃO na primeira coluna.

    O progresso é opcional e executado na thread de trabalho.
    """

    def _p(pct, msg):
        if progresso:
            progresso(pct, msg)

    _p(0, "Montando planilhas…")

    t0 = time.perf_counter()

    wb = Workbook()

    ws_tt = wb.active
    ws_tt.title = "TimeTable"

    _montar_aba(
        ws_tt,
        df_tt,
        bg_tt
    )

    ws_cen = wb.create_sheet("Cenario")

    _montar_aba(
        ws_cen,
        df_cen,
        bg_cen,
        fontes_cen
    )

    print(
        f"[tempo] montagem: {time.perf_counter() - t0:.1f}s"
    )

    t1 = time.perf_counter()

    relatorio = processar_em_memoria(
        wb,
        wb,
        "Cenario",
        "TimeTable",
        progresso=lambda p, m: _p(10 + p * 0.8, m),
    )

    print(
        f"[tempo] regras: {time.perf_counter() - t1:.1f}s"
    )

    _p(92, "Lendo resultado…")

    t2 = time.perf_counter()

    # Resultado completo, incluindo TIPO DE ALTERAÇÃO.
    df_completo, bgs_completo, fontes_completo = (
        _ler_aba_processada(wb["Cenario"])
    )

    # Separa a versão visual da versão de exportação.
    (
        df_visual,
        bgs_visual,
        fontes_visual,
        df_export,
        bgs_export,
        fontes_export
    ) = _separar_coluna_tipo_alteracao(
        df_completo,
        bgs_completo,
        fontes_completo
    )

    print(
        f"[tempo] leitura: {time.perf_counter() - t2:.1f}s | "
        f"{len(df_completo)} linhas x {df_completo.shape[1]} colunas"
    )

    wb.close()

    _p(100, "Concluído")

    return ResultadoProcessamento(
        df=df_visual,
        bgs=bgs_visual,
        fontes=fontes_visual,

        df_export=df_export,
        bgs_export=bgs_export,
        fontes_export=fontes_export,

        relatorio=relatorio
    )


# ============================================================
# MODELO DA TABELA
# ============================================================

class PlanilhaModel(QAbstractTableModel):
    """Modelo virtual: a tabela só pinta as células visíveis."""

    def __init__(self, df=None, bgs=None, fontes=None, cores_cabecalho=None, fontes_cabecalho=None):
        super().__init__()

        self.definir(
            df,
            bgs,
            fontes,
            cores_cabecalho,
            fontes_cabecalho
        )

    def definir(self, df, bgs=None, fontes=None, cores_cabecalho=None, fontes_cabecalho=None):

        self.beginResetModel()

        self.df = df
        self.bgs = bgs or []
        self.fontes = fontes or []
        self.cores_cabecalho = list(cores_cabecalho or [])
        self.fontes_cabecalho = list(fontes_cabecalho or [])

        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):

        if parent.isValid() or self.df is None:
            return 0

        return len(self.df)

    def columnCount(self, parent=QModelIndex()):

        if parent.isValid() or self.df is None:
            return 0

        return self.df.shape[1]

    def headerData(
        self,
        section,
        orientation,
        role=Qt.DisplayRole
    ):
        """Fornece texto e cores do cabeçalho original do arquivo Excel."""
        if self.df is None:
            return None

        if orientation == Qt.Horizontal:
            if role == Qt.DisplayRole:
                return str(self.df.columns[section])

            if role == Qt.BackgroundRole and section < len(self.cores_cabecalho):
                cor = self.cores_cabecalho[section]
                if cor:
                    return QBrush(QColor(cor))

            if role == Qt.ForegroundRole and section < len(self.fontes_cabecalho):
                cor = self.fontes_cabecalho[section]
                if cor:
                    return QBrush(QColor(cor))

            if role == Qt.TextAlignmentRole:
                return Qt.AlignCenter

            return None

        if orientation == Qt.Vertical and role == Qt.DisplayRole:
            return str(section + 1)

        return None

    def data(self, index, role=Qt.DisplayRole):

        if not index.isValid() or self.df is None:
            return None

        r = index.row()
        c = index.column()

        if role == Qt.DisplayRole:

            valor = self.df.iat[r, c]

            if valor is None:
                return ""

            try:
                if valor != valor:
                    return ""
            except Exception:
                pass

            return str(valor)

        if role == Qt.BackgroundRole:

            if r < len(self.bgs) and c < len(self.bgs[r]):
                return QColor(self.bgs[r][c])

        if role == Qt.ForegroundRole:

            if (
                r < len(self.fontes)
                and c < len(self.fontes[r])
                and self.fontes[r][c]
            ):
                return QColor(self.fontes[r][c])

            # Evita texto branco sobre branco em tema escuro.
            return QColor("#000000")

        if role == Qt.TextAlignmentRole:
            return Qt.AlignCenter

        return None


# ============================================================
# WORKER
# ============================================================

class Worker(QObject):

    concluido = Signal(object)
    falhou = Signal(str)
    progresso = Signal(int, str)

    def __init__(self, fn, usa_progresso=False):

        super().__init__()

        self.fn = fn
        self.usa_progresso = usa_progresso

    def emitir(self, pct, msg=""):
        self.progresso.emit(
            int(pct),
            str(msg)
        )

    @Slot()
    def run(self):

        try:

            resultado = (
                self.fn(self.emitir)
                if self.usa_progresso
                else self.fn()
            )

        except Exception as exc:

            traceback.print_exc()

            self.falhou.emit(
                f"{type(exc).__name__}: {exc}"
            )

            return

        self.concluido.emit(resultado)


# ============================================================
# EXECUÇÃO EM THREAD SEPARADA
# ============================================================

class Execucao(QObject):
    """
    Executa fn em uma QThread e chama on_ok / on_err / on_progress
    na thread da interface.
    """

    _ativas: set = set()

    def __init__(
        self,
        fn,
        on_ok=None,
        on_err=None,
        on_progress=None
    ):

        super().__init__()

        self._on_ok = on_ok
        self._on_err = on_err
        self._on_progress = on_progress

        self._thread = QThread()

        self._worker = Worker(
            fn,
            usa_progresso=on_progress is not None
        )

        self._worker.moveToThread(
            self._thread
        )

        self._thread.started.connect(
            self._worker.run
        )

        self._worker.concluido.connect(
            self._receber_ok
        )

        self._worker.falhou.connect(
            self._receber_erro
        )

        self._worker.progresso.connect(
            self._receber_progresso
        )

        self._thread.finished.connect(
            self._worker.deleteLater
        )

        self._thread.finished.connect(
            self._finalizar
        )

        # Mantém a referência viva até terminar.
        Execucao._ativas.add(self)

    def iniciar(self):
        self._thread.start()

    @Slot(int, str)
    def _receber_progresso(self, pct, msg):

        if self._on_progress:
            self._on_progress(pct, msg)

    @Slot(object)
    def _receber_ok(self, payload):

        self._thread.quit()

        if self._on_ok:
            self._on_ok(payload)

    @Slot(str)
    def _receber_erro(self, msg):

        self._thread.quit()

        if self._on_err:
            self._on_err(msg)

    @Slot()
    def _finalizar(self):

        self._thread.wait()
        self._thread.deleteLater()

        Execucao._ativas.discard(self)

        self.deleteLater()

    @classmethod
    def aguardar_todas(cls):

        for execucao in list(cls._ativas):
            execucao._thread.wait()


# ============================================================
# INTEGRAÇÃO OPT+Z (CONSULTA E EXPORTAÇÃO OFICIAL)
# ============================================================

OPTZ_URL_LOGIN = "http://optz2.integra.itg:5002/signin"
OPTZ_URL_API = "http://optz2.integra.itg:5000"
OPTZ_URL_EXPORTACOES = "http://optz2.integra.itg:5005"
OPTZ_CONTA = "jca"
OPTZ_TIMEOUT = 30
OPTZ_TIMEOUT_DOWNLOAD = 120
OPTZ_INTERVALO_CONSULTA = 5
OPTZ_TEMPO_MAXIMO_ESPERA = 1500  # 25 minutos
OPTZ_MODELO_EXPORTACAO = 1


def _optz_json(resposta, contexto):
    try:
        dados = resposta.json()
    except ValueError as erro:
        raise ValueError(
            f"{contexto}: resposta da API não é JSON válido: "
            f"{(resposta.text or '')[:500]}"
        ) from erro
    if not isinstance(dados, dict):
        raise ValueError(f"{contexto}: formato de resposta inesperado.")
    return dados


def _optz_autenticar(session, usuario, senha):
    resposta = session.post(
        OPTZ_URL_LOGIN,
        json={"conta": OPTZ_CONTA, "usuario": usuario, "senha": senha},
        headers={"Accept": "application/json"},
        timeout=OPTZ_TIMEOUT,
    )
    resposta.raise_for_status()
    retorno = _optz_json(resposta, "Autenticação OPT+Z")
    token = (retorno.get("data") or {}).get("accessToken")
    if not token:
        raise ValueError("O OPT+Z não retornou accessToken no login.")
    session.headers.update({"Authorization": f"Bearer {token}", "Accept": "application/json"})


def _optz_listar_planejamentos(session):
    itens = []
    pagina = 1
    while True:
        resposta = session.get(
            f"{OPTZ_URL_API}/planejamentos",
            params={"limit": 50, "page": pagina, "pesquisa": "", "sort": "nome:asc"},
            timeout=OPTZ_TIMEOUT,
        )
        resposta.raise_for_status()
        retorno = _optz_json(resposta, "Consulta de planejamentos")
        registros = retorno.get("data", [])
        if not isinstance(registros, list):
            raise ValueError("A API retornou uma lista de planejamentos inválida.")
        if not registros:
            break
        itens.extend(registros)
        total = retorno.get("total")
        if (total is not None and len(itens) >= int(total)) or len(registros) < 50:
            break
        pagina += 1
    if not itens:
        raise ValueError("Nenhum planejamento foi encontrado no OPT+Z.")
    return itens


def _optz_ultimo_cenario_ho(session, planejamento_id):
    pagina = 1
    while True:
        resposta = session.get(
            f"{OPTZ_URL_API}/cenarios",
            params={
                "limit": 20, "page": pagina, "pesquisa": "",
                "planejamentoId": planejamento_id, "sort": "criadoEm:desc",
            },
            timeout=OPTZ_TIMEOUT,
        )
        resposta.raise_for_status()
        retorno = _optz_json(resposta, "Consulta de cenários")
        cenarios = retorno.get("data", [])
        if not isinstance(cenarios, list):
            raise ValueError("A API retornou uma lista de cenários inválida.")
        if not cenarios:
            break
        # Ordenação criadaEm:desc: o primeiro código terminado em HO é o mais recente.
        for cenario in cenarios:
            codigo = str(cenario.get("codigo") or "").strip()
            if codigo.upper().endswith("HO"):
                if not cenario.get("id"):
                    raise ValueError("O cenário HO encontrado não possui identificador.")
                return cenario
        total = retorno.get("total")
        if (total is not None and pagina * 20 >= int(total)) or len(cenarios) < 20:
            break
        pagina += 1
    raise ValueError("Não foi encontrado cenário cujo código termine em HO nesse planejamento.")


def _optz_solicitar_exportacao(session, cenario_id):
    resposta = session.post(
        f"{OPTZ_URL_EXPORTACOES}/exportacoes/entidade/cenario",
        params={"confirmed": "false"},
        json={"cenarioId": cenario_id, "modelo": OPTZ_MODELO_EXPORTACAO},
        timeout=OPTZ_TIMEOUT,
    )
    resposta.raise_for_status()
    retorno = _optz_json(resposta, "Solicitação de exportação")
    dados = retorno.get("data") or {}
    exportacao_id = dados.get("id") or dados.get("exportacaoId")
    if not exportacao_id:
        raise ValueError(f"A API não retornou o ID da exportação. Campos: {list(dados)}")
    return str(exportacao_id)


def _optz_baixar_ate_concluir(session, exportacao_id, codigo_cenario, progresso=None):
    url = f"{OPTZ_URL_EXPORTACOES}/exportacoes/{exportacao_id}"
    inicio = time.monotonic()
    ultimo_erro = "arquivo ainda indisponível"
    while time.monotonic() - inicio <= OPTZ_TEMPO_MAXIMO_ESPERA:
        resposta = session.get(
            url,
            headers={"Accept": "application/octet-stream, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, */*"},
            timeout=OPTZ_TIMEOUT_DOWNLOAD,
        )
        if resposta.status_code == 200 and resposta.content:
            conteudo = resposta.content
            try:
                with zipfile.ZipFile(io.BytesIO(conteudo)) as zf:
                    nomes = set(zf.namelist())
                    if "xl/workbook.xml" not in nomes:
                        raise ValueError("A resposta não contém uma pasta de trabalho XLSX válida.")
                    corrompido = zf.testzip()
                    if corrompido:
                        raise ValueError(f"Arquivo XLSX corrompido: {corrompido}")
            except zipfile.BadZipFile:
                trecho = conteudo[:500].decode("utf-8", errors="replace")
                ultimo_erro = trecho or "resposta não é um XLSX"
            else:
                pasta = Path.cwd() / "exportacoes_optz"
                pasta.mkdir(parents=True, exist_ok=True)
                nome = "".join(c if c.isalnum() or c in "-_" else "_" for c in str(codigo_cenario or "CENARIO_HO"))
                destino = pasta / f"CENARIO_{nome}.xlsx"
                destino.write_bytes(conteudo)
                return destino
        else:
            try:
                erro_json = resposta.json()
                erros = erro_json.get("errors", []) if isinstance(erro_json, dict) else []
                ultimo_erro = "; ".join(map(str, erros)) if isinstance(erros, list) else str(erros)
            except ValueError:
                ultimo_erro = (resposta.text or f"HTTP {resposta.status_code}")[:500]
            # A resposta 400 com 'Exportação não concluída' é esperada durante o processamento.
            normalizado = unicodedata.normalize("NFKD", ultimo_erro).encode("ascii", "ignore").decode().lower()
            if resposta.status_code == 400 and not any(
                termo in normalizado
                for termo in ("exportacao nao concluida", "arquivo nao disponivel", "ainda processando")
            ):
                resposta.raise_for_status()
            if resposta.status_code not in (202, 400, 404, 409, 425) and resposta.status_code >= 400:
                resposta.raise_for_status()
        decorrido = int(time.monotonic() - inicio)
        if progresso:
            progresso(min(95, int(decorrido * 95 / OPTZ_TEMPO_MAXIMO_ESPERA)),
                      f"Aguardando geração do arquivo no OPT+Z ({decorrido // 60} min)…")
        time.sleep(OPTZ_INTERVALO_CONSULTA)
    raise TimeoutError(
        f"O OPT+Z não disponibilizou o XLSX em 25 minutos. Última resposta: {ultimo_erro}"
    )


def importar_cenario_optz(usuario, senha, planejamento_id, progresso=None):
    """Solicita exportação oficial do cenário HO mais recente e carrega o XLSX."""
    with requests.Session() as session:
        if progresso:
            progresso(2, "Autenticando no OPT+Z…")
        _optz_autenticar(session, usuario, senha)
        if progresso:
            progresso(5, "Localizando o cenário HO mais recente…")
        cenario = _optz_ultimo_cenario_ho(session, planejamento_id)
        if progresso:
            progresso(8, f"Solicitando exportação de {cenario.get('codigo')}…")
        exportacao_id = _optz_solicitar_exportacao(session, cenario["id"])
        caminho = _optz_baixar_ate_concluir(
            session, exportacao_id, cenario.get("codigo") or "HO", progresso
        )
    if progresso:
        progresso(97, "Importando planilha exportada…")
    dados = carregar_cenario(str(caminho))
    dados["caminho"] = str(caminho)
    _aplicar_estilos_cabecalho(dados, str(caminho))
    dados["origem_optz"] = True
    dados["cenario_optz"] = cenario
    return dados


# ============================================================
# FILTRO POR COLUNA (ESTILO EXCEL / GOOGLE SHEETS)
# ============================================================

ESPACO_ICONE_FILTRO = 26


def _texto_celula(valor):
    """Mesmo texto que a tabela exibe na célula."""
    if valor is None:
        return ""

    try:
        if valor != valor:  # NaN
            return ""
    except Exception:
        pass

    return str(valor)


def _chave_ordenacao(texto):
    """Números primeiro (em ordem numérica), depois textos, vazias por último."""
    if texto == "":
        return (2, 0.0, "")

    try:
        return (0, float(texto), "")
    except ValueError:
        return (1, 0.0, _normalizar_texto(texto))


class FiltroLinhas(QSortFilterProxyModel):
    """Filtro por coluna: cada coluna aceita um conjunto de valores marcados.

    Filtros em várias colunas funcionam como E. Os dados originais não mudam;
    apenas as linhas são ocultadas na visualização.
    """

    def __init__(self, parent=None):
        super().__init__(parent)

        self._selecoes = {}   # {coluna: set(textos permitidos)}
        self._mascaras = {}   # {coluna: [bool por linha]}
        self._mascara = None  # combinação (E) de todas as colunas
        self._textos = {}     # cache: {coluna: [texto por linha]}

    def setSourceModel(self, modelo):
        super().setSourceModel(modelo)
        # Dados novos invalidam filtros e cache.
        modelo.modelAboutToBeReset.connect(self._limpar_tudo)

    def _limpar_tudo(self):
        self._selecoes = {}
        self._mascaras = {}
        self._mascara = None
        self._textos = {}

    def _df(self):
        modelo = self.sourceModel()
        return None if modelo is None else modelo.df

    def textos_coluna(self, coluna):
        textos = self._textos.get(coluna)

        if textos is None:
            textos = [
                _texto_celula(v)
                for v in self._df().iloc[:, coluna].tolist()
            ]
            self._textos[coluna] = textos

        return textos

    def coluna_filtrada(self, coluna):
        return coluna in self._selecoes

    def colunas_filtradas(self):
        return sorted(self._selecoes)

    def selecao(self, coluna):
        return self._selecoes.get(coluna)

    def valores_disponiveis(self, coluna):
        """Contagem dos valores da coluna nas linhas que passam pelos
        filtros das OUTRAS colunas (como no Excel)."""
        textos = self.textos_coluna(coluna)

        outras = [
            m for c, m in self._mascaras.items()
            if c != coluna
        ]

        if outras:
            visiveis = [all(v) for v in zip(*outras)]
            return Counter(
                t for t, ok in zip(textos, visiveis) if ok
            )

        return Counter(textos)

    def definir_selecao(self, coluna, selecionados):
        """selecionados: conjunto de textos permitidos; None remove o filtro."""
        if selecionados is None:
            self._selecoes.pop(coluna, None)
            self._mascaras.pop(coluna, None)
        else:
            permitidos = set(selecionados)
            self._selecoes[coluna] = permitidos
            self._mascaras[coluna] = [
                t in permitidos
                for t in self.textos_coluna(coluna)
            ]

        self._recalcular()

    def limpar_filtros(self):
        self._selecoes = {}
        self._mascaras = {}
        self._recalcular()

    def _recalcular(self):
        if not self._mascaras:
            self._mascara = None
        else:
            self._mascara = [
                all(v) for v in zip(*self._mascaras.values())
            ]

        self.invalidateFilter()

    def filterAcceptsRow(self, linha_origem, pai):
        mascara = self._mascara

        if mascara is None or linha_origem >= len(mascara):
            return True

        return mascara[linha_origem]


class _CheckSelecionarTudo(QCheckBox):
    """Caixa de três estados: o clique alterna só entre marcado e desmarcado."""

    def nextCheckState(self):
        self.setCheckState(
            Qt.Unchecked
            if self.checkState() == Qt.Checked
            else Qt.Checked
        )


class PopupFiltroColuna(QWidget):
    """Janela de filtro de uma coluna: pesquisa, selecionar tudo e valores."""

    # Conjunto de textos marcados, ou None para limpar o filtro da coluna.
    aplicado = Signal(object)

    def __init__(self, parent, valores, selecionados=None):
        super().__init__(parent, Qt.Popup | Qt.FramelessWindowHint)

        self.setAttribute(Qt.WA_DeleteOnClose, True)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setFixedWidth(300)

        self._itens = []  # (item, texto original, texto normalizado)

        cartao = QFrame(self)
        cartao.setObjectName("cartaoFiltro")

        self.campo_busca = QLineEdit()
        self.campo_busca.setPlaceholderText("Pesquisar…")
        self.campo_busca.setClearButtonEnabled(True)

        self.check_todos = _CheckSelecionarTudo("(Selecionar tudo)")
        self.check_todos.setTristate(True)

        self.lista = QListWidget()
        self.lista.setUniformItemSizes(True)
        self.lista.setSelectionMode(QListWidget.NoSelection)
        self.lista.setMinimumHeight(240)

        for texto in valores:
            exibicao = texto if texto != "" else "(Vazias)"

            item = QListWidgetItem(exibicao)
            item.setFlags(Qt.ItemIsEnabled)
            item.setCheckState(
                Qt.Checked
                if selecionados is None or texto in selecionados
                else Qt.Unchecked
            )

            self.lista.addItem(item)
            self._itens.append(
                (item, texto, _normalizar_texto(exibicao))
            )

        botao_limpar = QPushButton("Limpar")
        botao_cancelar = QPushButton("Cancelar")
        botao_ok = QPushButton("OK")
        botao_ok.setDefault(True)

        linha_botoes = QHBoxLayout()
        linha_botoes.addWidget(botao_limpar)
        linha_botoes.addStretch(1)
        linha_botoes.addWidget(botao_cancelar)
        linha_botoes.addWidget(botao_ok)

        interno = QVBoxLayout(cartao)
        interno.setContentsMargins(10, 10, 10, 10)
        interno.setSpacing(8)
        interno.addWidget(self.campo_busca)
        interno.addWidget(self.check_todos)
        interno.addWidget(self.lista, 1)
        interno.addLayout(linha_botoes)

        externo = QVBoxLayout(self)
        externo.setContentsMargins(0, 0, 0, 0)
        externo.addWidget(cartao)

        self.setStyleSheet("""
            QFrame#cartaoFiltro {
                background-color: rgba(255, 255, 255, 240);
                border: 1px solid rgba(0, 0, 0, 40);
                border-radius: 10px;
            }
            QListWidget {
                background-color: #FFFFFF;
                border: 1px solid #D9E1E8;
                border-radius: 6px;
                color: #263445;
                outline: 0;
            }
            QCheckBox {
                background: transparent;
                color: #263445;
            }
            QPushButton {
                padding: 5px 14px;
                min-width: 60px;
                min-height: 16px;
            }
        """)

        # Pesquisa com pequena pausa para listas grandes.
        self._timer_busca = QTimer(self)
        self._timer_busca.setSingleShot(True)
        self._timer_busca.setInterval(150)
        self._timer_busca.timeout.connect(self._filtrar_lista)

        self.campo_busca.textChanged.connect(
            lambda _t: self._timer_busca.start()
        )
        self.campo_busca.returnPressed.connect(self._confirmar)
        self.lista.itemClicked.connect(self._ao_clicar_item)
        self.check_todos.clicked.connect(self._ao_clicar_todos)

        botao_ok.clicked.connect(self._confirmar)
        botao_cancelar.clicked.connect(self.close)
        botao_limpar.clicked.connect(self._limpar)

        self._atualizar_selecionar_tudo()

    def _ao_clicar_item(self, item):
        item.setCheckState(
            Qt.Unchecked
            if item.checkState() == Qt.Checked
            else Qt.Checked
        )
        self._atualizar_selecionar_tudo()

    def _ao_clicar_todos(self):
        estado = (
            Qt.Checked
            if self.check_todos.checkState() == Qt.Checked
            else Qt.Unchecked
        )

        # Age só sobre os itens visíveis (resultado da pesquisa).
        for item, _texto, _norm in self._itens:
            if not item.isHidden():
                item.setCheckState(estado)

        self._atualizar_selecionar_tudo()

    def _atualizar_selecionar_tudo(self):
        visiveis = [
            item for item, _t, _n in self._itens
            if not item.isHidden()
        ]

        marcados = sum(
            1 for item in visiveis
            if item.checkState() == Qt.Checked
        )

        if visiveis and marcados == len(visiveis):
            estado = Qt.Checked
        elif marcados == 0:
            estado = Qt.Unchecked
        else:
            estado = Qt.PartiallyChecked

        self.check_todos.setCheckState(estado)

    def _filtrar_lista(self):
        termo = _normalizar_texto(self.campo_busca.text().strip())

        self.lista.setUpdatesEnabled(False)

        for item, _texto, norm in self._itens:
            item.setHidden(bool(termo) and termo not in norm)

        self.lista.setUpdatesEnabled(True)

        self._atualizar_selecionar_tudo()

    def _confirmar(self):
        marcados = {
            texto
            for item, texto, _n in self._itens
            if item.checkState() == Qt.Checked
        }

        self.aplicado.emit(marcados)
        self.close()

    def _limpar(self):
        self.aplicado.emit(None)
        self.close()


class CabecalhoColorido(QHeaderView):
    """Cabeçalho com a cor do Excel e botão de filtro em cada coluna.

    O fundo e o texto são desenhados diretamente (o QHeaderView costuma
    ignorar o BackgroundRole quando há QSS), e o botão de filtro fica no
    canto direito de cada seção.
    """

    filtro_clicado = Signal(int)

    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)

        self._coluna_filtrada = lambda coluna: False
        self._ignorar_soltar = False

        self.setMouseTracking(True)
        self.setHighlightSections(False)

    def definir_funcao_filtrada(self, funcao):
        self._coluna_filtrada = funcao

    # --- geometria do botão -------------------------------------

    def _rect_secao(self, coluna):
        return QRect(
            self.sectionViewportPosition(coluna),
            0,
            self.sectionSize(coluna),
            self.height(),
        )

    @staticmethod
    def _rect_icone(rect):
        return QRect(rect.right() - 22, rect.center().y() - 9, 18, 18)

    def _sobre_icone(self, pos):
        coluna = self.logicalIndexAt(pos)

        if coluna < 0:
            return -1

        area = self._rect_icone(self._rect_secao(coluna)).adjusted(-3, -3, 3, 3)

        return coluna if area.contains(pos) else -1

    # --- mouse --------------------------------------------------

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            coluna = self._sobre_icone(event.position().toPoint())

            if coluna >= 0:
                self._ignorar_soltar = True
                event.accept()
                self.filtro_clicado.emit(coluna)
                return

        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        if self._ignorar_soltar:
            self._ignorar_soltar = False
            event.accept()
            return

        super().mouseReleaseEvent(event)

    def mouseMoveEvent(self, event):
        if self._sobre_icone(event.position().toPoint()) >= 0:
            self.viewport().setCursor(Qt.PointingHandCursor)
        else:
            self.viewport().unsetCursor()

        super().mouseMoveEvent(event)

    # --- desenho ------------------------------------------------

    def paintSection(self, painter, rect, logicalIndex):

        modelo = self.model()

        if modelo is None or not rect.isValid():
            super().paintSection(painter, rect, logicalIndex)
            return

        fundo = modelo.headerData(
            logicalIndex, self.orientation(), Qt.BackgroundRole
        )

        if not isinstance(fundo, QBrush):
            super().paintSection(painter, rect, logicalIndex)
            self._desenhar_icone(painter, rect, logicalIndex, QColor("#EEEEEE"))
            return

        frente = modelo.headerData(
            logicalIndex, self.orientation(), Qt.ForegroundRole
        )

        cor_texto = (
            frente.color()
            if isinstance(frente, QBrush)
            else QColor("#000000")
        )

        texto = str(
            modelo.headerData(
                logicalIndex, self.orientation(), Qt.DisplayRole
            ) or ""
        )

        painter.save()

        painter.fillRect(rect, fundo)

        painter.setPen(QPen(QColor(0, 0, 0, 60)))
        painter.drawRect(rect.adjusted(0, 0, -1, -1))

        area = rect.adjusted(6, 0, -ESPACO_ICONE_FILTRO, 0)

        painter.setPen(cor_texto)
        painter.drawText(
            area,
            Qt.AlignCenter,
            QFontMetrics(self.font()).elidedText(
                texto, Qt.ElideRight, max(0, area.width())
            ),
        )

        painter.restore()

        self._desenhar_icone(painter, rect, logicalIndex, fundo.color())

    def _desenhar_icone(self, painter, rect, coluna, cor_base):

        if rect.width() < 34:
            return

        caixa = QRectF(self._rect_icone(rect))
        centro = caixa.center()

        filtrada = bool(self._coluna_filtrada(coluna))

        claro = (
            0.299 * cor_base.red()
            + 0.587 * cor_base.green()
            + 0.114 * cor_base.blue()
        ) > 140

        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setPen(Qt.NoPen)

        if filtrada:
            # Coluna filtrada: botão destacado e ícone de funil.
            painter.setBrush(QColor(118, 177, 178, 240))
            painter.drawRoundedRect(caixa, 4, 4)

            painter.setBrush(QColor("#0E3862"))
            painter.drawPolygon(QPolygonF([
                QPointF(centro.x() - 5, centro.y() - 4),
                QPointF(centro.x() + 5, centro.y() - 4),
                QPointF(centro.x() + 1, centro.y() + 0.5),
                QPointF(centro.x() + 1, centro.y() + 5),
                QPointF(centro.x() - 1, centro.y() + 3.5),
                QPointF(centro.x() - 1, centro.y() + 0.5),
            ]))
        else:
            painter.setBrush(
                QColor(255, 255, 255, 120 if claro else 60)
            )
            painter.drawRoundedRect(caixa, 4, 4)

            painter.setBrush(
                QColor("#263445") if claro else QColor("#FFFFFF")
            )
            painter.drawPolygon(QPolygonF([
                QPointF(centro.x() - 4, centro.y() - 2),
                QPointF(centro.x() + 4, centro.y() - 2),
                QPointF(centro.x(), centro.y() + 3),
            ]))

        painter.restore()


# ============================================================
# DELEGATE DAS CÉLULAS (COR DE FUNDO E DE FONTE DO EXCEL)
# ============================================================

class DelegateCelulaColorida(QStyledItemDelegate):
    """Pinta fundo e texto das células com as cores do modelo.

    Com o QSS global (QTableView::item), o Qt pode ignorar BackgroundRole e
    ForegroundRole. Este delegate desenha as células diretamente, então a cor
    exibida é sempre a lida do Excel.
    """

    def paint(self, painter, option, index):

        fundo = index.data(Qt.BackgroundRole)
        frente = index.data(Qt.ForegroundRole)

        if isinstance(fundo, QBrush):
            fundo = fundo.color()

        if isinstance(frente, QBrush):
            frente = frente.color()

        painter.save()

        if fundo is not None:
            painter.fillRect(option.rect, QColor(fundo))

        # Seleção como véu translúcido: a cor do Excel continua visível.
        if option.state & QStyle.State_Selected:
            painter.fillRect(option.rect, QColor(118, 177, 178, 90))

        texto = index.data(Qt.DisplayRole)

        if texto:
            area = option.rect.adjusted(6, 0, -6, 0)

            painter.setFont(option.font)
            painter.setPen(
                QColor(frente) if frente is not None else QColor("#000000")
            )
            painter.drawText(
                area,
                Qt.AlignCenter,
                option.fontMetrics.elidedText(
                    str(texto), Qt.ElideRight, max(0, area.width())
                ),
            )

        painter.restore()


# ============================================================
# ABA IMPORTADOR
# ============================================================

class AbaImportador(QWidget):
    """Importa Excel: o cabeçalho é localizado automaticamente."""

    importado = Signal()
    importar_optz_solicitado = Signal()

    def __init__(self, texto_botao: str, tipo: str):

        super().__init__()

        self.tipo = tipo

        self.caminho_arquivo = None
        self.planilha_encontrada = None
        self.header_row = 2

        self.df = None
        self.bgs = None
        self.fontes = None
        self.cores_cabecalho = []
        self.fontes_cabecalho = []

        # Cópia dos dados originais.
        # Aplicar alterações sempre parte daqui.
        self.df_original = None
        self.bgs_original = None
        self.fontes_original = None

        self.cols = None

        self.botao_importar = QPushButton(
            texto_botao
        )

        self.botao_importar.clicked.connect(
            self.escolher_arquivo
        )

        self.botao_optz = None
        if self.tipo == "cen":
            self.botao_optz = QPushButton("Importar do OPT+Z")
            self.botao_optz.clicked.connect(lambda _checked=False: self.importar_optz_solicitado.emit())

        self.label_arquivo = QLabel(
            "Nenhum arquivo selecionado"
        )

        self.label_arquivo.setWordWrap(True)

        self.modelo = PlanilhaModel()

        # A tabela exibe o modelo filtrado (proxy), sem alterar os dados.
        self.filtro = FiltroLinhas(self)
        self.filtro.setSourceModel(self.modelo)

        self.tabela = QTableView()

        cabecalho_colorido = CabecalhoColorido(Qt.Horizontal, self.tabela)
        cabecalho_colorido.definir_funcao_filtrada(
            self.filtro.coluna_filtrada
        )
        cabecalho_colorido.filtro_clicado.connect(
            self._abrir_filtro_coluna
        )

        self.tabela.setHorizontalHeader(cabecalho_colorido)

        self.tabela.setModel(
            self.filtro
        )

        self.tabela.setItemDelegate(
            DelegateCelulaColorida(self.tabela)
        )

        self.label_filtros = QLabel("")
        self.label_filtros.setWordWrap(True)

        self.botao_limpar_filtros = QPushButton("Limpar filtros")
        self.botao_limpar_filtros.setEnabled(False)

        self.botao_limpar_filtros.clicked.connect(
            self._limpar_filtros
        )

        self._popup_filtro = None

        self.tabela.setEditTriggers(
            QTableView.NoEditTriggers
        )

        self.tabela.setAlternatingRowColors(False)

        self.tabela.verticalHeader().setDefaultSectionSize(22)
        # Mantém a faixa dos números das linhas branca mesmo ao selecionar uma linha.
        self.tabela.verticalHeader().setHighlightSections(False)

        header = self.tabela.horizontalHeader()

        header.setSectionResizeMode(
            QHeaderView.Interactive
        )

        header.setStretchLastSection(False)

        layout = QVBoxLayout(self)

        layout.addWidget(self.botao_importar)
        if self.botao_optz is not None:
            layout.addWidget(self.botao_optz)
        layout.addWidget(self.label_arquivo)

        linha_filtro = QHBoxLayout()
        linha_filtro.addWidget(self.label_filtros, 1)
        linha_filtro.addWidget(self.botao_limpar_filtros)

        layout.addLayout(linha_filtro)
        layout.addWidget(self.tabela)

    def escolher_arquivo(self):

        # .xls não é suportado pelo openpyxl.
        caminho, _ = QFileDialog.getOpenFileName(
            self,
            "Selecione um arquivo Excel",
            "",
            "Arquivos Excel (*.xlsx *.xlsm);;Todos os arquivos (*)",
        )

        if not caminho:
            return

        self._importar(caminho)

    def _importar(self, caminho: str):

        self.label_arquivo.setText(
            f"Lendo {Path(caminho).name}…"
        )

        def tarefa():

            # O arquivo é lido diretamente.
            if self.tipo == "tt":

                dados = carregar_timetable(
                    caminho
                )

            elif self.tipo == "cen":

                dados = carregar_cenario(
                    caminho
                )

            else:

                raise ValueError(
                    f"Tipo de planilha inválido: {self.tipo}"
                )

            dados["caminho"] = caminho
            _aplicar_estilos_cabecalho(dados, caminho)

            return dados

        self._rodar(
            tarefa,
            self._aplicar_carga,
            self._falha_carga
        )

    def _falha_carga(self, msg: str):

        self.label_arquivo.setText(
            "Falha na importação."
        )

        QMessageBox.critical(
            self,
            "Erro",
            f"Não foi possível abrir o arquivo:\n{msg}"
        )

    def _aplicar_carga(self, dados: dict):

        self.caminho_arquivo = dados["caminho"]
        self.planilha_encontrada = dados["aba"]
        self.header_row = dados["header_row"]

        self.df = dados["df"]
        self.bgs = dados["bgs"]
        self.fontes = dados["fontes"]
        # Mesma regra de cor das células: #RRGGBB.
        self.cores_cabecalho = [
            _normalizar_cor(c, "#FFFFFF")
            for c in dados.get("header_bgs", [])
        ]
        self.fontes_cabecalho = [
            _normalizar_cor(c, "#000000")
            for c in dados.get("header_fontes", [])
        ]

        self.df_original = self.df
        self.bgs_original = self.bgs
        self.fontes_original = self.fontes

        self.cols = dados["cols"]

        n = (
            0
            if self.df is None
            else len(self.df)
        )

        self.label_arquivo.setText(
            f"{Path(self.caminho_arquivo).name}  |  "
            f"aba {self.planilha_encontrada}  |  "
            f"{n} linhas (cabeçalho na linha {self.header_row})"
        )

        self.modelo.definir(
            self.df,
            self.bgs,
            self.fontes,
            self.cores_cabecalho,
            self.fontes_cabecalho,
        )

        self._ajustar_colunas()

        self._resetar_filtros()

        self.importado.emit()

    def atualizar(self, df, bgs, fontes):
        """Mostra o resultado processado, mantendo a cor original de cada cabeçalho."""

        self.df = df
        self.bgs = bgs
        self.fontes = fontes

        # Associa os estilos do cabeçalho original aos nomes das colunas
        # visíveis. Isso também funciona quando TIPO DE ALTERAÇÃO é removida
        # da visualização após o processamento.
        indices_originais = {
            str(nome): i
            for i, nome in enumerate(self.df_original.columns)
        } if self.df_original is not None else {}

        cores_visiveis = []
        fontes_visiveis = []
        for nome in df.columns:
            indice = indices_originais.get(str(nome))
            if indice is not None and indice < len(self.cores_cabecalho):
                cores_visiveis.append(self.cores_cabecalho[indice])
            else:
                cores_visiveis.append("#FFFFFF")

            if indice is not None and indice < len(self.fontes_cabecalho):
                fontes_visiveis.append(self.fontes_cabecalho[indice])
            else:
                fontes_visiveis.append("#000000")

        self.modelo.definir(
            df,
            bgs,
            fontes,
            cores_visiveis,
            fontes_visiveis,
        )

        self._ajustar_colunas()

        self._resetar_filtros()

    # --------------------------------------------------------
    # AJUSTE DE COLUNAS E FILTRO POR COLUNA
    # --------------------------------------------------------

    def _ajustar_colunas(self):
        """Ajusta ao conteúdo e reserva espaço para o botão de filtro."""

        self.tabela.resizeColumnsToContents()

        cabecalho = self.tabela.horizontalHeader()

        for i in range(cabecalho.count()):
            cabecalho.resizeSection(
                i,
                cabecalho.sectionSize(i) + ESPACO_ICONE_FILTRO
            )

    def _resetar_filtros(self):
        """Remove todos os filtros (dados novos ou reprocessados)."""

        self.filtro.limpar_filtros()
        self._atualizar_resumo_filtro()

    def _abrir_filtro_coluna(self, coluna):
        """Abre a lista de valores da coluna, como no Excel."""

        if (
            self.df is None
            or coluna < 0
            or coluna >= self.modelo.columnCount()
        ):
            return

        disponiveis = self.filtro.valores_disponiveis(coluna)
        valores = sorted(disponiveis, key=_chave_ordenacao)

        popup = PopupFiltroColuna(
            self,
            valores,
            self.filtro.selecao(coluna),
        )

        popup.aplicado.connect(
            lambda selecionados, c=coluna:
                self._aplicar_selecao_coluna(c, selecionados)
        )

        popup.adjustSize()

        cabecalho = self.tabela.horizontalHeader()

        ponto = cabecalho.mapToGlobal(
            QPoint(
                cabecalho.sectionViewportPosition(coluna),
                cabecalho.height(),
            )
        )

        tela = (
            QApplication.screenAt(ponto)
            or QApplication.primaryScreen()
        )
        area = tela.availableGeometry()

        x = max(area.left(), min(ponto.x(), area.right() - popup.width()))
        y = ponto.y()

        if y + popup.height() > area.bottom():
            y = max(
                area.top(),
                ponto.y() - cabecalho.height() - popup.height()
            )

        popup.move(x, y)

        # Guarda a referência enquanto o popup estiver aberto.
        self._popup_filtro = popup

        popup.show()
        popup.campo_busca.setFocus()

    def _aplicar_selecao_coluna(self, coluna, selecionados):

        if selecionados is not None:
            # Se todos os valores disponíveis estão marcados, não há filtro.
            disponiveis = set(self.filtro.valores_disponiveis(coluna))

            if disponiveis <= selecionados:
                selecionados = None

        self.filtro.definir_selecao(coluna, selecionados)

        self._atualizar_resumo_filtro()

    def _limpar_filtros(self):

        self.filtro.limpar_filtros()

        self._atualizar_resumo_filtro()

    def _atualizar_resumo_filtro(self):

        ativas = self.filtro.colunas_filtradas()

        self.botao_limpar_filtros.setEnabled(bool(ativas))

        # Redesenha os botões de filtro do cabeçalho.
        self.tabela.horizontalHeader().viewport().update()

        if not ativas or self.df is None:
            self.label_filtros.setText("")
            return

        nomes = ", ".join(
            str(self.df.columns[c])
            for c in ativas
            if c < len(self.df.columns)
        )

        self.label_filtros.setText(
            f"{self.filtro.rowCount()} de {self.modelo.rowCount()} linhas"
            f"  |  Filtrado por: {nomes}"
        )

    def _rodar(self, fn, on_ok, on_err):

        QApplication.setOverrideCursor(
            Qt.WaitCursor
        )

        self.botao_importar.setEnabled(False)

        def restaurar():

            QApplication.restoreOverrideCursor()

            self.botao_importar.setEnabled(True)

        def ok(payload):

            restaurar()

            on_ok(payload)

        def err(msg):

            restaurar()

            on_err(msg)

        Execucao(
            fn,
            ok,
            err
        ).iniciar()


# ============================================================
# JANELA PRINCIPAL
# ============================================================

class JanelaPrincipal(QWidget):

    def __init__(self):

        super().__init__()

        self.setWindowTitle(
            "Alteração Cenário Rodoviário"
        )
        self.setWindowIcon(QIcon(str(CAMINHO_ICONE)))

        self.resize(1100, 700)

        self.resultado: ResultadoProcessamento | None = None

        self.aba_timetable = AbaImportador(
            "Importar TimeTable",
            "tt"
        )

        self.aba_cenario = AbaImportador(
            "Importar Cenário",
            "cen"
        )

        self.aba_cenario.importar_optz_solicitado.connect(
            self.importar_cenario_optz
        )

        # Nova importação invalida o resultado anterior.
        self.aba_timetable.importado.connect(
            self._invalidar_resultado
        )

        self.aba_cenario.importado.connect(
            self._invalidar_resultado
        )

        self.botao_aplicar = QPushButton(
            "Aplicar alterações"
        )

        self.botao_aplicar.clicked.connect(
            self.aplicar_alteracoes
        )

        self.botao_baixar = QPushButton(
            "Baixar cenário processado"
        )

        self.botao_baixar.setEnabled(False)

        self.botao_baixar.clicked.connect(
            self.baixar_cenario
        )

        self.label_status = QLabel("")

        self.label_status.setWordWrap(True)

        self.abas = QTabWidget()

        self.abas.addTab(
            self.aba_timetable,
            "TimeTable"
        )

        self.abas.addTab(
            self.aba_cenario,
            "Cenário"
        )

        layout = QVBoxLayout(self)

        layout.addWidget(self.abas)
        layout.addWidget(self.botao_aplicar)
        layout.addWidget(self.botao_baixar)
        layout.addWidget(self.label_status)

    def closeEvent(self, event):

        Execucao.aguardar_todas()

        super().closeEvent(event)

    def _invalidar_resultado(self):

        self.resultado = None

        self.botao_baixar.setEnabled(False)

        self.label_status.setText("")

    def _rodar(
        self,
        fn,
        on_ok,
        on_err=None,
        botao_bloquear=None,
        on_progress=None
    ):

        if botao_bloquear is None:
            botao_bloquear = self.botao_aplicar

        QApplication.setOverrideCursor(
            Qt.WaitCursor
        )

        botao_bloquear.setEnabled(False)

        def restaurar_interface():

            QApplication.restoreOverrideCursor()

            botao_bloquear.setEnabled(True)

        def ok(payload):

            restaurar_interface()

            if on_ok:
                on_ok(payload)

        def err(msg):

            restaurar_interface()

            if on_err:

                on_err(msg)

            else:

                QMessageBox.critical(
                    self,
                    "Erro",
                    msg
                )

        Execucao(
            fn,
            ok,
            err,
            on_progress
        ).iniciar()

    # --------------------------------------------------------
    # IMPORTAR CENÁRIO DIRETAMENTE DO OPT+Z
    # --------------------------------------------------------

    def importar_cenario_optz(self):
        usuario, ok = QInputDialog.getText(
            self, "Conectar ao OPT+Z", "Usuário OPT+Z:"
        )
        if not ok or not usuario.strip():
            return
        senha, ok = QInputDialog.getText(
            self, "Conectar ao OPT+Z", "Senha OPT+Z:", QLineEdit.Password
        )
        if not ok or not senha:
            return

        self.label_status.setText("Conectando ao OPT+Z e consultando planejamentos…")
        self.aba_cenario.botao_optz.setEnabled(False)

        def buscar():
            with requests.Session() as session:
                _optz_autenticar(session, usuario.strip(), senha)
                return _optz_listar_planejamentos(session)

        def sucesso_planejamentos(planejamentos):
            self.aba_cenario.botao_optz.setEnabled(True)
            opcoes = []
            for item in planejamentos:
                empresa = (item.get("empresa") or {}).get("codigo") or ""
                opcoes.append(
                    f"{len(opcoes) + 1:02d} - {item.get('nome') or 'Sem nome'} | "
                    f"Empresa: {empresa} | Cenários: {item.get('numCenarios', 'N/I')}"
                )
            escolha, ok = QInputDialog.getItem(
                self, "Selecionar planejamento OPT+Z",
                "Escolha o planejamento:", opcoes, 0, False
            )
            if not ok or not escolha:
                self.label_status.setText("Importação do OPT+Z cancelada.")
                return
            indice = opcoes.index(escolha)
            planejamento = planejamentos[indice]
            planejamento_id = planejamento.get("id")
            if not planejamento_id:
                QMessageBox.critical(self, "OPT+Z", "O planejamento selecionado não possui ID.")
                return
            self._iniciar_importacao_optz(usuario.strip(), senha, planejamento_id, planejamento)

        def erro_planejamentos(msg):
            self.aba_cenario.botao_optz.setEnabled(True)
            self.label_status.setText("Falha ao consultar planejamentos no OPT+Z.")
            QMessageBox.critical(self, "Erro OPT+Z", msg)

        self._rodar(
            buscar, sucesso_planejamentos, erro_planejamentos,
            botao_bloquear=self.aba_cenario.botao_optz
        )

    def _iniciar_importacao_optz(self, usuario, senha, planejamento_id, planejamento):
        dlg = DialogoProgressoLinha(
            self,
            "Importando cenário",
            "Iniciando importação…",
        )
        dlg.show()

        def tarefa(progresso):
            return importar_cenario_optz(
                usuario, senha, planejamento_id, progresso=progresso
            )

        def progresso(pct, mensagem):
            dlg.setValue(pct)
            dlg.setLabelText(f"{mensagem} ({pct}%)")

        def sucesso(dados):
            dlg.setValue(100)
            dlg.close()
            self.aba_cenario._aplicar_carga(dados)
            self.abas.setCurrentWidget(self.aba_cenario)
            self.label_status.setText(
                f"Cenário importado do OPT+Z | Planejamento: {planejamento.get('nome')} | "
                f"Cenário: {(dados.get('cenario_optz') or {}).get('codigo')} | "
                f"Arquivo: {dados.get('caminho')}"
            )
            QMessageBox.information(
                self, "Importação concluída",
                "O cenário foi exportado pelo OPT+Z e importado para o Python.\n\n"
                f"Planejamento: {planejamento.get('nome')}\n"
                f"Cenário: {(dados.get('cenario_optz') or {}).get('codigo')}\n"
                f"Arquivo: {dados.get('caminho')}\n\n"
                "Nenhum dado do cenário foi alterado na plataforma."
            )

        def erro(mensagem):
            dlg.close()
            self.label_status.setText("Falha ao importar cenário do OPT+Z.")
            QMessageBox.critical(self, "Erro na importação OPT+Z", mensagem)

        self._rodar(
            tarefa, sucesso, erro,
            botao_bloquear=self.aba_cenario.botao_optz,
            on_progress=progresso,
        )

    # --------------------------------------------------------
    # APLICAR ALTERAÇÕES
    # --------------------------------------------------------

    def aplicar_alteracoes(self):

        tt = self.aba_timetable
        cen = self.aba_cenario

        if tt.df_original is None:

            QMessageBox.warning(
                self,
                "TimeTable não importado",
                "Importe o TimeTable antes de aplicar as alterações."
            )

            return

        if cen.df_original is None:

            QMessageBox.warning(
                self,
                "Cenário não importado",
                "Importe o Cenário antes de aplicar as alterações."
            )

            return

        self.label_status.setText(
            "Processando cenário…"
        )

        # Sempre a partir dos dados ORIGINAIS importados.
        df_tt = tt.df_original
        bg_tt = tt.bgs_original

        df_cen = cen.df_original
        bg_cen = cen.bgs_original
        fontes_cen = cen.fontes_original

        cols_tt = tt.cols
        cols_cen = cen.cols

        dlg = DialogoProgressoLinha(
            self,
            "Processando cenário",
            "Iniciando…",
        )
        dlg.show()

        def tarefa(progresso):

            return processar_cenario(
                df_tt,
                bg_tt,
                df_cen,
                bg_cen,
                cols_tt,
                cols_cen,
                fontes_cen,
                progresso=progresso,
            )

        def ao_progresso(pct, msg):

            dlg.setValue(pct)

            dlg.setLabelText(
                f"{msg}  ({pct}%)"
            )

        def sucesso(resultado):

            dlg.setValue(100)
            dlg.close()

            self._on_processado(resultado)

        def erro(mensagem):

            dlg.close()

            self.label_status.setText(
                "Falha ao processar o cenário."
            )

            QMessageBox.critical(
                self,
                "Erro ao processar",
                mensagem
            )

        self._rodar(
            tarefa,
            sucesso,
            erro,
            on_progress=ao_progresso
        )

    # --------------------------------------------------------
    # EXIBIR RESULTADO NO APLICATIVO
    # --------------------------------------------------------

    def _on_processado(self, resultado: ResultadoProcessamento):

        self.resultado = resultado

        # Exibe apenas o DataFrame sem TIPO DE ALTERAÇÃO.
        self.aba_cenario.atualizar(
            resultado.df,
            resultado.bgs,
            resultado.fontes
        )

        self.abas.setCurrentWidget(
            self.aba_cenario
        )

        self.botao_baixar.setEnabled(True)

        self.label_status.setText(
            resultado.relatorio.replace(
                "\n",
                "  |  "
            )
        )

    # --------------------------------------------------------
    # EXPORTAR EXCEL
    # --------------------------------------------------------

    def baixar_cenario(self):

        if self.resultado is None:

            QMessageBox.warning(
                self,
                "Atenção",
                "Não existe cenário processado para salvar.",
            )

            return

        origem = Path(
            self.aba_cenario.caminho_arquivo or "Cenario.xlsx"
        )

        sugerido = str(
            origem.with_name(
                origem.stem + "_processado.xlsx"
            )
        )

        caminho, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar cenário processado",
            sugerido,
            "Arquivos Excel (*.xlsx)",
        )

        if not caminho:
            return

        if not caminho.lower().endswith(".xlsx"):
            caminho += ".xlsx"

        destino = Path(
            caminho
        ).expanduser().resolve()

        resultado = self.resultado

        # IMPORTANTE:
        # A exportação utiliza os dados completos.
        # TIPO DE ALTERAÇÃO é a primeira coluna.
        kwargs = {
            "caminho_origem": self.aba_cenario.caminho_arquivo,
            "nome_aba": (
                self.aba_cenario.planilha_encontrada
                or "Cenario"
            ),
            "header_row": (
                self.aba_cenario.header_row
                or 1
            ),

            "df": resultado.df_export,
            "bgs": resultado.bgs_export,
            "fontes": resultado.fontes_export,

            "caminho_saida": str(destino),
        }

        self.label_status.setText(
            "Salvando arquivo..."
        )

        def tarefa():

            salvar_cenario(**kwargs)

            if not destino.is_file():

                raise FileNotFoundError(
                    "O arquivo não foi encontrado após a gravação: "
                    f"{destino}"
                )

            if destino.stat().st_size == 0:

                raise IOError(
                    "O arquivo salvo está vazio."
                )

            return str(destino)

        def sucesso(saida):

            self.label_status.setText(
                f"Arquivo salvo: {saida}"
            )

            QMessageBox.information(
                self,
                "Salvamento concluído",
                f"O arquivo foi salvo com sucesso em:\n{saida}",
            )

        def erro(mensagem):

            self.label_status.setText(
                "Falha ao salvar o arquivo."
            )

            QMessageBox.critical(
                self,
                "Erro ao salvar",
                f"Não foi possível salvar o cenário:\n\n{mensagem}",
            )

        self._rodar(
            tarefa,
            sucesso,
            erro,
            botao_bloquear=self.botao_baixar
        )


# ============================================================
# INICIALIZAÇÃO
# ============================================================

# Aplicação Flask
web_app = Flask(__name__)


# Inicialização da aplicação desktop
if __name__ == "__main__":
    qt_app = QApplication(sys.argv)
    qt_app.setWindowIcon(QIcon(str(CAMINHO_ICONE)))
    aplicar_estilo(qt_app)
    janela = JanelaPrincipal()
    aplicar_fundo_blur(janela)
    janela.show()
    sys.exit(qt_app.exec())
