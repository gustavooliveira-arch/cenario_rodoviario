"""Leitura e gravação de planilhas Excel com valores e cores de célula."""

from __future__ import annotations

import colorsys
import re
from copy import copy
from datetime import datetime, time, timedelta

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, PatternFill
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.utils import get_column_letter

from regras_cenario import (
    MAX_LINHAS_CABECALHO,
    NOMES_CEN,
    NOMES_TT,
    eh_erro,
    norm_cabecalho,
    norm_sem_acento,
    texto_bruto,
)

# Ordem das cores de tema usada pelo Excel nos índices "theme" das células:
# 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..9=accent1..6, 10=hlink, 11=folHlink
_TEMA_PADRAO = [
    "FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31",
    "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72",
]
_NOMES_TEMA = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3",
               "accent4", "accent5", "accent6", "hlink", "folHlink"]

_RE_HORA = re.compile(r"^\s*(\d{1,2}):(\d{2})\s*$")


def paleta_tema(wb) -> list[str]:
    """Lê as cores de tema do próprio arquivo (cai na paleta padrão do Office)."""
    xml = getattr(wb, "loaded_theme", None)
    if not xml:
        return _TEMA_PADRAO
    if isinstance(xml, bytes):
        xml = xml.decode("utf-8", errors="ignore")
    saida = []
    for nome in _NOMES_TEMA:
        m = re.search(
            rf"<a:{nome}>\s*<a:(?:srgbClr val|sysClr[^>]*?lastClr)=\"([0-9A-Fa-f]{{6}})\"",
            xml,
        )
        if not m:
            return _TEMA_PADRAO
        saida.append(m.group(1).upper())
    return saida


def _aplicar_tint(hex6: str, tint: float | None) -> str:
    """Clareia/escurece a cor como o Excel faz (tint entre -1 e 1)."""
    if not tint:
        return hex6
    r, g, b = (int(hex6[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, min(max(l, 0.0), 1.0), s)
    return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


def _hex_de_argb(valor) -> str | None:
    """
    Converte 'AARRGGBB' / 'RRGGBB' em '#RRGGBB'.

    O canal alfa é IGNORADO: o Excel não o utiliza e o openpyxl grava
    cores de 6 dígitos com prefixo '00' (e a tabela COLOR_INDEX também usa
    '00'). Tratar '00' como "transparente" descartava quase todas as cores.
    """
    s = str(valor).upper().lstrip("#")
    if len(s) == 8:
        s = s[2:]
    if len(s) != 6:
        return None
    try:
        int(s, 16)
    except ValueError:
        return None
    return "#" + s


def cor_para_hex(color, tema: list[str] | None = None) -> str | None:
    """Converte cores RGB, tema e índice do Excel para #RRGGBB."""
    if color is None:
        return None

    tipo = getattr(color, "type", None)

    if tipo == "rgb":
        rgb = getattr(color, "rgb", None)
        if not rgb:
            return None
        # "00000000" é o valor padrão do openpyxl para "sem cor definida"
        if str(rgb).upper() == "00000000":
            return None
        return _hex_de_argb(rgb)

    if tipo == "theme":
        idx = getattr(color, "theme", None)
        paleta = tema or _TEMA_PADRAO
        if idx is None:
            return None
        try:
            idx = int(idx)
        except (TypeError, ValueError):
            return None
        if not 0 <= idx < len(paleta):
            return None
        return "#" + _aplicar_tint(
            paleta[idx], getattr(color, "tint", 0) or 0
        )

    if tipo == "indexed":
        idx = getattr(color, "indexed", None)
        if idx is None:
            return None
        try:
            idx = int(idx)
        except (TypeError, ValueError):
            return None
        # 64/65 = cores de sistema (primeiro plano / fundo automáticos)
        if not 0 <= idx < min(64, len(COLOR_INDEX)):
            return None
        return _hex_de_argb(COLOR_INDEX[idx])

    return None


def cor_preenchimento(cell, tema: list[str] | None = None) -> str:
    fill = cell.fill
    if fill is None or fill.fill_type in (None, "none"):
        return "#FFFFFF"
    for candidato in (fill.fgColor, fill.bgColor):
        hex_ = cor_para_hex(candidato, tema)
        if hex_:
            return hex_
    return "#FFFFFF"


def cor_fonte(cell, tema: list[str] | None = None) -> str | None:
    font = cell.font
    if font is None or font.color is None:
        return None
    hex_ = cor_para_hex(font.color, tema)
    # preto é a cor padrão: não precisa ser gravado
    if hex_ in (None, "#000000"):
        return None
    return hex_


def _luminancia(hex_cor: str) -> float:
    s = hex_cor.replace("#", "")[:6]
    r, g, b = (int(s[i:i + 2], 16) for i in (0, 2, 4))
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255


def _claro(hex_cor: str | None) -> bool:
    return bool(hex_cor) and _luminancia(hex_cor) > 0.85


def normalizar_hora_exibicao(valor) -> object:
    if valor is None:
        return ""
    if isinstance(valor, (datetime, time)):
        return valor.strftime("%H:%M")
    if isinstance(valor, timedelta):
        total = int(valor.total_seconds() // 60)
        return f"{(total // 60) % 24:02d}:{total % 60:02d}"
    texto = str(valor)
    m = re.match(r"^\s*(\d{1,2}):(\d{2})", texto)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return valor


def _coluna_eh_hora(col_nome: str) -> bool:
    nome = norm_sem_acento(col_nome)
    return "PARTIDA" in nome or "CHEGADA" in nome


def _valor_celula(v, eh_hora: bool) -> object:
    if v is None or eh_erro(v):
        return ""
    if eh_hora:
        if isinstance(v, (datetime, time, timedelta)):
            return normalizar_hora_exibicao(v)
        if isinstance(v, (int, float)) and not isinstance(v, bool) and 0 <= float(v) < 1.5:
            total = int(round(float(v) * 24 * 60))
            return f"{(total // 60) % 24:02d}:{total % 60:02d}"
    return v


def localizar_cabecalho_ws(ws, nomes: list[str]) -> tuple[int, list[int]]:
    """Localiza todos os cabeçalhos na mesma linha (a que tiver mais nomes esperados)."""
    alvos = [norm_cabecalho(n) for n in nomes]
    ult_col = ws.max_column or 1
    ult_lin = min(ws.max_row or 1, MAX_LINHAS_CABECALHO)

    melhor_linha, melhores, melhor_pont = 0, {}, 0
    for r, linha in enumerate(
            ws.iter_rows(min_row=1, max_row=ult_lin, max_col=ult_col, values_only=True), 1):
        pos: dict[str, int] = {}
        for c, v in enumerate(linha, 1):
            t = norm_cabecalho(v)
            if t in alvos and t not in pos:
                pos[t] = c
        if len(pos) > melhor_pont:
            melhor_linha, melhores, melhor_pont = r, pos, len(pos)
        if melhor_pont == len(alvos):
            break

    faltando = [n for n, a in zip(nomes, alvos) if a not in melhores]
    if faltando:
        raise ValueError(
            f"Não encontrei todos os cabeçalhos na mesma linha da aba '{ws.title}'.\n"
            f"Melhor linha encontrada: {melhor_linha}\n"
            f"Faltando: {', '.join(faltando)}"
        )
    return melhor_linha, [melhores[a] for a in alvos]


def _nome_aba(wb, nomes_aceitaveis: tuple[str, ...]) -> str:
    alvos = [norm_sem_acento(n) for n in nomes_aceitaveis]
    nomes = list(wb.sheetnames)
    for n in nomes:
        if norm_sem_acento(n) in alvos:
            return n
    for n in nomes:
        nn = norm_sem_acento(n)
        if any(a in nn or nn.startswith(a) for a in alvos):
            return n
    raise ValueError(
        f"O arquivo não possui uma aba chamada '{nomes_aceitaveis[0]}'.\n"
        f"Abas disponíveis: {', '.join(nomes)}"
    )


def _ultima_linha_util(ws, colunas_1based: list[int], lin_cab: int) -> int:
    ult = ws.max_row or lin_cab
    if ult <= lin_cab:
        return lin_cab
    for r in range(ult, lin_cab, -1):
        if any(texto_bruto(ws.cell(r, c).value) for c in colunas_1based):
            return r
    return lin_cab


def carregar_planilha(
    caminho: str,
    nomes_aba: tuple[str, ...],
    nomes_cabecalho: list[str],
) -> dict:
    """
    Lê a aba em uma única passada pelo openpyxl: valores, cor de fundo e
    cor da fonte de cada célula ficam sempre alinhados (mesma linha/coluna).

    Antes, os valores vinham do pandas (segunda abertura do arquivo, com
    skiprows/nrows) e as cores do openpyxl, o que duplicava o custo de
    leitura e podia desalinhar dados e cores.
    """
    wb = load_workbook(caminho, data_only=True)
    try:
        tema = paleta_tema(wb)
        nome_aba = _nome_aba(wb, nomes_aba)
        ws = wb[nome_aba]

        lin_cab, cols_obrig = localizar_cabecalho_ws(ws, nomes_cabecalho)
        ult_col = ws.max_column or 1

        # Títulos: nomes vazios viram Unnamed_N e duplicados recebem sufixo.
        titulos: list[str] = []
        usados: dict[str, int] = {}
        valores_cabecalho = next(
            ws.iter_rows(
                min_row=lin_cab,
                max_row=lin_cab,
                max_col=ult_col,
                values_only=True,
            )
        )
        for c, valor in enumerate(valores_cabecalho, start=1):
            bruto = texto_bruto(valor)
            nome = bruto if bruto else f"Unnamed_{c}"
            if nome in usados:
                usados[nome] += 1
                nome = f"{nome}_{usados[nome]}"
            else:
                usados[nome] = 1
            titulos.append(nome)

        eh_cenario = [norm_cabecalho(n) for n in nomes_cabecalho] == [
            norm_cabecalho(n) for n in NOMES_CEN
        ]
        if eh_cenario:
            colunas_limite = [c for c in cols_obrig[:3] if c]
        else:
            # LINHA 1, SERVIÇO 1 e TIPO DE ALTERAÇÃO
            colunas_limite = [cols_obrig[0], cols_obrig[2], cols_obrig[6]]

        ult_lin = _ultima_linha_util(ws, colunas_limite, lin_cab)
        quantidade = max(0, ult_lin - lin_cab)

        eh_hora = [_coluna_eh_hora(t) for t in titulos]
        linhas: list[list] = []
        bgs: list[list[str]] = []
        fontes: list[list[str | None]] = []

        if quantidade:
            for row in ws.iter_rows(
                min_row=lin_cab + 1,
                max_row=ult_lin,
                max_col=ult_col,
            ):
                vals, linha_bg, linha_ft = [], [], []
                for c, cell in enumerate(row):
                    vals.append(_valor_celula(cell.value, eh_hora[c]))
                    bg = cor_preenchimento(cell, tema)
                    ft = cor_fonte(cell, tema)
                    if ft and _claro(ft) and _claro(bg):
                        ft = None
                    linha_bg.append(bg)
                    linha_ft.append(ft)
                linhas.append(vals)
                bgs.append(linha_bg)
                fontes.append(linha_ft)

        if linhas:
            df = pd.DataFrame(linhas, columns=titulos, dtype=object)
        else:
            df = pd.DataFrame(columns=titulos)

        mapa_cols = []
        for nome in nomes_cabecalho:
            alvo = norm_cabecalho(nome)
            idx = next((i for i, t in enumerate(titulos) if norm_cabecalho(t) == alvo), -1)
            mapa_cols.append(idx)
    finally:
        wb.close()

    return {
        "df": df,
        "bgs": bgs,
        "fontes": fontes,
        "caminho": caminho,
        "aba": nome_aba,
        "header_row": lin_cab,
        "cols": mapa_cols,
        "colunas": titulos,
    }


def carregar_timetable(caminho: str) -> dict:
    return carregar_planilha(caminho, ("TimeTable",), NOMES_TT)


def carregar_cenario(caminho: str) -> dict:
    return carregar_planilha(caminho, ("Cenario", "Cenário"), NOMES_CEN)


def _fill_hex(hex_cor: str) -> PatternFill:
    rgb = (hex_cor or "#FFFFFF").replace("#", "")[:6]
    return PatternFill(fill_type="solid", fgColor="FF" + rgb)


def salvar_cenario(
    caminho_origem: str | None,
    nome_aba: str,
    header_row: int,
    df: pd.DataFrame,
    bgs: list[list[str]],
    fontes: list[list[str | None]],
    caminho_saida: str,
) -> None:
    """Grava o Cenário processado (valores + cores) e mantém as demais abas."""
    if caminho_origem:
        wb = load_workbook(caminho_origem)
        if nome_aba in wb.sheetnames:
            ws = wb[nome_aba]
        else:
            ws = wb.create_sheet(nome_aba)
            for c, nome in enumerate(df.columns, start=1):
                ws.cell(1, c, str(nome))
            header_row = 1
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = nome_aba or "Cenario"
        for c, nome in enumerate(df.columns, start=1):
            ws.cell(1, c, str(nome))
        header_row = 1

    if ws.max_row > header_row:
        ws.delete_rows(header_row + 1, ws.max_row - header_row)

    n = len(df)
    ncols = len(df.columns)
    eh_col_hora = [
        norm_sem_acento(nome) in ("PARTIDA", "CHEGADA") for nome in df.columns
    ]

    for i in range(n):
        r = header_row + 1 + i
        for c in range(ncols):
            valor = df.iat[i, c]
            if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
                valor = None
            elif isinstance(valor, str) and valor == "":
                valor = None

            # Horários "HH:MM" voltam a ser horários reais do Excel
            # (antes eram gravados como texto, com number_format sem efeito).
            if eh_col_hora[c] and isinstance(valor, str):
                m = _RE_HORA.match(valor)
                if m and int(m.group(1)) < 24 and int(m.group(2)) < 60:
                    valor = time(int(m.group(1)), int(m.group(2)))

            cell = ws.cell(r, c + 1, valor)

            bg = bgs[i][c] if i < len(bgs) and c < len(bgs[i]) else "#FFFFFF"
            ft = fontes[i][c] if i < len(fontes) and c < len(fontes[i]) else None

            # Sempre sobrescreve o preenchimento. Não fazer isso deixa cores
            # antigas no arquivo quando a nova célula não possui cor.
            cell.fill = _fill_hex(bg or "#FFFFFF")

            # Mantém negrito, tamanho, família e demais atributos da fonte,
            # alterando somente a cor.
            if ft and _claro(ft) and _claro(bg):
                ft = None
            fonte_atual = copy(cell.font)
            fonte_atual.color = "FF" + ft.replace("#", "")[:6] if ft else "FF000000"
            cell.font = fonte_atual

            cell.alignment = Alignment(horizontal="center", vertical="center")
            if isinstance(valor, time):
                cell.number_format = "hh:mm"

    for c in range(1, ncols + 1):
        letra = get_column_letter(c)
        if ws.column_dimensions[letra].width is None:
            ws.column_dimensions[letra].width = 12

    wb.save(caminho_saida)