#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conferência de alterações: TimeTable -> aba de Cenário (arquivo .xlsx)
Conversão em Python do Google Apps Script "ProcessarCenario".

Uso:
    python processar_cenario.py arquivo.xlsx "Cenário 1"
    python processar_cenario.py arquivo.xlsx "Cenário 1" -o resultado.xlsx

Requisitos:  pip install openpyxl

Premissas (iguais às do script original):
  - O arquivo tem uma aba "TimeTable" e a aba de cenário informada.
  - As duas abas têm uma coluna com cabeçalho "CHAVE"; o cruzamento é feito só por ela.
  - Os valores da coluna CHAVE precisam estar gravados no arquivo (valor em cache).
    Se a CHAVE for fórmula, abra o arquivo no Excel e salve antes de rodar.

O arquivo original NUNCA é sobrescrito: o resultado vai para um arquivo novo.
"""

import argparse
import colorsys
import copy
import datetime as dt
import re
import sys
import unicodedata
from bisect import bisect_right
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import PatternFill
from openpyxl.styles.colors import COLOR_INDEX, Color
from openpyxl.utils import get_column_letter

ABA_TIMETABLE = "TimeTable"
MAX_LINHAS_CABECALHO = 30

COR = {
    "VERDE_CLARO": "#C6EFCE",
    "VERMELHO_CLARO": "#FFC7CE",
    "VERMELHO_ESCURO": "#EE1E1E",
    "VINHO_SERVICO": "#A61C00",
    "VERDE_INCLUSAO_SERVICO": "#00B050",
    "VERDE_SERVICO_CELULA": "#93C47D",
    "LARANJA": "#FFBBA2",
    "AMARELO": "#FFFF00",
    "AMARELO_CLARO": "#FFFF99",
    "VERMELHO": "#FF0000",
    "AZUL_CLARO": "#ADC9EE",
}

NOMES_TT = ["LINHA 1", "LINHA 2", "SERVIÇO 1", "SERVIÇO 2", "ETAPA", "TIPO DE HORÁRIO",
            "TIPO DE ALTERAÇÃO", "ORIGEM", "DESTINO", "HR PARTIDA", "HR. CHEGADA", "FROTA",
            "SEG", "TER", "QUA", "QUI", "SEX", "SAB", "DOM"]
NOMES_CEN = ["LINHA", "SERVIÇO", "ETAPA/SEÇÃO", "TIPO",
             "ORIGEM", "DESTINO", "PARTIDA", "CHEGADA", "FROTA",
             "SEG", "TER", "QUA", "QUI", "SEX", "SÁB", "DOM"]
IDX_CHAVE_TT = len(NOMES_TT) - 1    # 19
IDX_CHAVE_CEN = len(NOMES_CEN) - 1  # 16


class ErroProcesso(Exception):
    pass


# ============================ TEXTO / VALORES ============================

_ERROS = re.compile(r"^#(N/A|REF!|VALUE!|DIV/0!|NAME\?|NUM!|NULL!|ERROR!)")


def eh_erro(v):
    return isinstance(v, str) and bool(_ERROS.match(v))


def _str(v):
    if isinstance(v, bool):
        return str(v).upper()
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v)


def texto_bruto(v):
    if v is None or eh_erro(v):
        return ""
    return _str(v).replace("\u00a0", " ").strip()


def norm(v):
    if v is None or eh_erro(v):
        return ""
    return _str(v).replace("\u00a0", " ").strip().upper()


def norm_sem_acento(v):
    s = norm(v)
    if any(ord(ch) > 127 for ch in s):
        s = unicodedata.normalize("NFD", s)
        s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip()


def valor_ou_branco(v):
    if v is None or eh_erro(v):
        return ""
    if texto_bruto(v) == "-":
        return ""
    return v


def _hhmm(h, m):
    return "%02d:%02d" % (h, m)


def ler_valor(v):
    """Horários (time/datetime/timedelta) viram texto 'HH:MM', como no script original."""
    if isinstance(v, (dt.datetime, dt.time)):
        return _hhmm(v.hour, v.minute)
    if isinstance(v, dt.timedelta):
        tot = int(v.total_seconds() // 60)
        return _hhmm(tot // 60, tot % 60)
    return v


def para_hora(v):
    """Converte 'HH:MM' em horário real do Excel (para gravar com formato hh:mm)."""
    if isinstance(v, str):
        m = re.match(r"^\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*$", v)
        if m and int(m.group(1)) < 24 and int(m.group(2)) < 60:
            return dt.time(int(m.group(1)), int(m.group(2)))
    elif isinstance(v, (int, float)) and not isinstance(v, bool) and 0 <= v < 1:
        s = round(v * 86400)
        return dt.time((s // 3600) % 24, (s % 3600) // 60)
    return v


# ============================ CORES ============================

_TEMA = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31",
         "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"]


def _tint(hex6, tint):
    r, g, b = (int(hex6[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, min(max(l, 0), 1), s)
    return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


def cor_hex(cell):
    """Cor de preenchimento da célula em '#RRGGBB' ('#FFFFFF' se não tiver preenchimento)."""
    f = cell.fill
    if f is None or f.fill_type != "solid":
        return "#FFFFFF"
    c = f.fgColor
    try:
        if c.type == "rgb" and isinstance(c.rgb, str):
            return "#" + c.rgb[-6:]
        if c.type == "theme" and c.theme is not None and 0 <= c.theme < len(_TEMA):
            return "#" + _tint(_TEMA[c.theme], c.tint or 0)
        if c.type == "indexed" and c.indexed is not None and c.indexed < len(COLOR_INDEX):
            return "#" + COLOR_INDEX[c.indexed][-6:]
    except Exception:
        pass
    return "#FFFFFF"


def _rgb(hexs):
    m = re.match(r"^#?([0-9a-f]{6})", hexs or "", re.I)
    if not m:
        return 255, 255, 255
    n = int(m.group(1), 16)
    return (n >> 16) & 255, (n >> 8) & 255, n & 255


def cor_eh_branca(hexs):
    if not hexs:
        return True
    r, g, b = _rgb(hexs)
    return r >= 250 and g >= 250 and b >= 250


def cor_eh_verde(hexs):
    r, g, b = _rgb(hexs)
    return g > r + 15 and g > b + 15


def cor_eh_amarelo(hexs):
    r, g, b = _rgb(hexs)
    return r >= 200 and g >= 200 and abs(r - g) <= 40 and b <= (r + g) // 2 - 40


def cor_eh_vermelho(hexs):
    r, g, b = _rgb(hexs)
    return r > g + 15 and r > b + 15


# ============================ CLASSIFICAÇÃO DOS TIPOS ============================

def eh_exclusao(tipo):
    t = norm_sem_acento(tipo)
    return (("EXCLUSAO" in t and "FREQ" not in t) or
            ("CANCELAMENTO" in t and "FREQ" not in t) or
            "INATIVACAO" in t)


def eh_exclusao_servico(tipo):
    t = norm_sem_acento(tipo)
    return "EXCLUSAO" in t and "SERVICO" in t and "FREQ" not in t


def eh_reducao_freq_e_alt_horario(tipo):
    t = norm_sem_acento(tipo)
    return "REDUCAO" in t and "FREQ" in t and "HORARIO" in t


def eh_reducao_frequencia(tipo):
    t = norm_sem_acento(tipo)
    if eh_reducao_freq_e_alt_horario(t):
        return False
    return "REDUCAO" in t or ("FREQ" in t and "EXCLUSAO" in t)


def eh_alteracao_frequencia(tipo):
    t = norm_sem_acento(tipo)
    if eh_reducao_freq_e_alt_horario(t):
        return False
    return "ALTERACAO" in t and "FREQ" in t


def eh_alteracao_horario(tipo):
    t = norm_sem_acento(tipo)
    return "ALTERACAO" in t and "HORARIO" in t and "FREQ" not in t


def eh_inclusao_frequencia(tipo):
    t = norm_sem_acento(tipo)
    return ("INCLUSAO" in t and "FREQ" in t) or ("INC" in t and "FREQ" in t)


def eh_inclusao(tipo):
    t = norm_sem_acento(tipo)
    if eh_inclusao_frequencia(t):
        return False
    return "INCLUSAO" in t


def eh_inclusao_servico(tipo):
    t = norm_sem_acento(tipo)
    return "INCLUSAO" in t and "SERVICO" in t and "FREQ" not in t


def eh_marcacao_s(v):
    return texto_bruto(v).upper() == "S"


def eh_marcacao_reducao(valor, bg):
    txt = texto_bruto(valor).upper()
    return txt in ("N", "X") or cor_eh_vermelho(bg)


# ============================ CHAVE / EXIBIÇÃO ============================
def norm_cabecalho(v):
    """Normaliza só para comparar cabeçalhos: sem acento, sem pontos, espaços/quebras unificados."""
    s = norm_sem_acento(v).replace(".", " ")
    return re.sub(r"\s+", " ", s).strip()

def norm_chave(v):
    s = re.sub(r"\s*/\s*", "/", norm_sem_acento(v))
    return "" if s == "-" else s


def combinar_exibicao(p1, p2):
    a, b = texto_bruto(p1), texto_bruto(p2)
    return a if b in ("-", "") else a + "/" + b


# ============================ CABEÇALHO / LIMITES ============================

def localizar_cabecalho(ws, nomes):
    """Devolve (linha_cabecalho, colunas 1-based de cada nome, nomes não encontrados)."""
    ult_lin = min(ws.max_row, MAX_LINHAS_CABECALHO)
    pos = {}
    for r, linha in enumerate(
            ws.iter_rows(min_row=1, max_row=ult_lin, max_col=ws.max_column, values_only=True), 1):
        for c, v in enumerate(linha, 1):
            t = norm_sem_acento(v)
            if t and t not in pos:
                pos[t] = (r, c)
    cols, linha, faltando = [0] * len(nomes), 0, []
    for k, nome in enumerate(nomes):
        p = pos.get(norm_sem_acento(nome))
        if p:
            cols[k] = p[1]
            linha = linha or p[0]
        else:
            faltando.append(nome)
    return linha, cols, faltando


def ultima_linha_util(ws, colunas, lin_cab):
    """Última linha com valor real em alguma das colunas (ignora linhas em branco/formatadas)."""
    maxr = lin_cab
    for c in colunas:
        for r in range(ws.max_row, lin_cab, -1):
            if texto_bruto(ws.cell(r, c).value):
                maxr = max(maxr, r)
                break
    return maxr


def calcular_ranks_dias(cols_cen):
    pos = [cols_cen[9 + d] for d in range(7)]
    rank = []
    for d in range(7):
        rk = 1
        for k in range(7):
            if pos[k] < pos[d] or (pos[k] == pos[d] and k < d):
                rk += 1
        rank.append(rk)
    return rank


# ============================ ESCRITA EM LOTE ============================

class Lote:
    """Acumula valores/cores/alinhamentos (a última escrita na célula vence) e grava de uma vez."""

    def __init__(self):
        self.limpar()

    def limpar(self):
        self.bg, self.ft, self.nf, self.vl = {}, {}, {}, {}
        self.ct = set()

    def pintar(self, r, c1, c2, cor):
        for c in range(c1, c2 + 1):
            self.bg[(r, c)] = cor

    def fonte(self, r, c1, c2, cor):
        for c in range(c1, c2 + 1):
            self.ft[(r, c)] = cor

    def centralizar(self, r, c1, c2):
        for c in range(c1, c2 + 1):
            self.ct.add((r, c))

    def formato_hora(self, r, c):
        self.nf[(r, c)] = "hh:mm"

    def valor(self, r, c, v):
        self.vl[(r, c)] = v

    def aplicar(self, ws):
        fills = {}
        for (r, c), cor in self.bg.items():
            if cor not in fills:
                fills[cor] = PatternFill("solid", fgColor="FF" + cor.lstrip("#").upper())
            ws.cell(r, c).fill = fills[cor]

        for (r, c), cor in self.ft.items():
            cell = ws.cell(r, c)
            f = copy.copy(cell.font)
            f.color = Color(rgb="FF" + cor.lstrip("#").upper())
            cell.font = f

        for (r, c) in self.ct:
            cell = ws.cell(r, c)
            a = copy.copy(cell.alignment)
            a.horizontal, a.vertical = "center", "center"
            cell.alignment = a

        for (r, c), fmt in self.nf.items():
            ws.cell(r, c).number_format = fmt

        for (r, c), v in self.vl.items():
            if (r, c) in self.nf:
                v = para_hora(v)
            ws.cell(r, c).value = None if v == "" else v


# ============================ INSERÇÃO DE LINHAS ============================

def inserir_linhas(ws, lista):
    """
    Abre espaço para todas as linhas novas de uma só vez (equivale ao insertDimension em lote).
    Reposiciona as células existentes e ajusta fórmulas de referência relativa da própria linha.
    Atenção: mesclagens, formatação condicional, validações e filtros da aba NÃO são deslocados.
    """
    alvos = sorted(it["targetRow"] for it in lista)
    novo = {}
    for (r, c), cell in ws._cells.items():
        nr = r + bisect_right(alvos, r)
        if nr != r:
            cell.row = nr
            if isinstance(cell.value, str) and cell.value.startswith("="):
                col = get_column_letter(c)
                cell.value = Translator(cell.value, origin="%s%d" % (col, r)) \
                    .translate_formula("%s%d" % (col, nr))
        novo[(nr, c)] = cell
    ws._cells = novo


# ============================ PROCESSO PRINCIPAL ============================

def executar(entrada, saida, nome_aba_cenario, nome_aba_tt=ABA_TIMETABLE):
    entrada = Path(entrada)
    wb = load_workbook(entrada, keep_vba=entrada.suffix.lower() == ".xlsm")  # fórmulas + estilos
    wbv = load_workbook(entrada, data_only=True)                              # valores em cache

    for nome in (nome_aba_tt, nome_aba_cenario):
        if nome not in wb.sheetnames:
            raise ErroProcesso("Aba '%s' não encontrada. Abas do arquivo: %s"
                               % (nome, ", ".join(wb.sheetnames)))

    ws_tt_v, ws_cen_v = wbv[nome_aba_tt], wbv[nome_aba_cenario]
    ws_cen = wb[nome_aba_cenario]

    lin_cab_tt, cols_tt, falt = localizar_cabecalho(ws_tt_v, NOMES_TT)
    if falt:
        raise ErroProcesso("Cabeçalhos não encontrados na aba '%s': %s" % (nome_aba_tt, ", ".join(falt)))
    lin_cab_cen, cols_cen, falt = localizar_cabecalho(ws_cen_v, NOMES_CEN)
    if falt:
        raise ErroProcesso("Cabeçalhos não encontrados na aba '%s': %s"
                           % (nome_aba_cenario, ", ".join(falt)))

    day_rank = calcular_ranks_dias(cols_cen)

    ult_lin_tt = ultima_linha_util(ws_tt_v, [cols_tt[0], cols_tt[2], cols_tt[6]], lin_cab_tt)
    ult_lin_cen = ultima_linha_util(ws_cen_v, [cols_cen[0], cols_cen[1], cols_cen[2]], lin_cab_cen)
    if ult_lin_tt <= lin_cab_tt or ult_lin_cen <= lin_cab_cen:
        raise ErroProcesso("Não foram encontrados dados abaixo do cabeçalho em alguma das abas.")
    ult_col_tt = ws_tt_v.max_column
    ult_col_cen = ws_cen.max_column

    # ---- Leitura: TimeTable (só linhas com "Tipo de Alteração") ----
    tt_dados, tt_bg = {}, {}
    for r in range(lin_cab_tt + 1, ult_lin_tt + 1):
        if texto_bruto(ws_tt_v.cell(r, cols_tt[6]).value):
            cells = [ws_tt_v.cell(r, c) for c in range(1, ult_col_tt + 1)]
            tt_dados[r] = [ler_valor(c.value) for c in cells]
            tt_bg[r] = [cor_hex(c) for c in cells]

    def tt_v(r, c):
        linha = tt_dados.get(r)
        if not linha or c - 1 >= len(linha) or linha[c - 1] is None:
            return ""
        return linha[c - 1]

    def tt_cor(r, c):
        linha = tt_bg.get(r)
        return linha[c - 1] if linha and c - 1 < len(linha) else "#FFFFFF"

    # ---- Leitura: Cenário (valores em cache) ----
    cen = {}
    for r, linha in enumerate(ws_cen_v.iter_rows(
            min_row=1, max_row=ult_lin_cen, max_col=ult_col_cen, values_only=True), 1):
        cen[r] = list(linha)

    def cen_v(r, c):
        linha = cen.get(r)
        if not linha or c - 1 >= len(linha) or linha[c - 1] is None:
            return ""
        return linha[c - 1]

    def chave_tt(r):
        partes = [combinar_exibicao(tt_v(r, cols_tt[0]), tt_v(r, cols_tt[1])),   # linha
                  combinar_exibicao(tt_v(r, cols_tt[2]), tt_v(r, cols_tt[3])),   # serviço
                  tt_v(r, cols_tt[4]),                                           # etapa
                  tt_v(r, cols_tt[7]), tt_v(r, cols_tt[8])]                      # origem, destino
        partes = [norm_chave(p) for p in partes]
        return "|".join(partes) if any(partes) else ""

    def chave_cen(r):
        partes = [cen_v(r, cols_cen[0]), cen_v(r, cols_cen[1]), cen_v(r, cols_cen[2]),
                  cen_v(r, cols_cen[4]), cen_v(r, cols_cen[5])]
        partes = [norm_chave(p) for p in partes]
        return "|".join(partes) if any(partes) else ""

    lote = Lote()

    def set_cen(r, c, v):
        cen[r][c - 1] = v
        lote.valor(r, c, v)

    def rank_linha(r):
        for d in range(7):
            if texto_bruto(cen_v(r, cols_cen[9 + d])).upper() == "S":
                return day_rank[d]
        return 0

    col_chave_cen = cols_cen[IDX_CHAVE_CEN]
    col_chave_tt = cols_tt[IDX_CHAVE_TT]

    # Se a CHAVE do cenário é fórmula linha a linha, as linhas novas recebem a mesma fórmula;
    # caso contrário recebem o valor da CHAVE da TimeTable.
    f_ini = ws_cen.cell(lin_cab_cen + 1, col_chave_cen).value
    f_fim = ws_cen.cell(ult_lin_cen, col_chave_cen).value
    modelo_formula = None
    if (isinstance(f_ini, str) and f_ini.startswith("=") and
            isinstance(f_fim, str) and f_fim.startswith("=")):
        modelo_formula = (f_ini, lin_cab_cen + 1)

    def chave_nova(r, linha_tt):
        if modelo_formula:
            col = get_column_letter(col_chave_cen)
            return Translator(modelo_formula[0], origin="%s%d" % (col, modelo_formula[1])) \
                .translate_formula("%s%d" % (col, r))
        return tt_v(linha_tt, col_chave_tt)

    q = dict(com_tipo=0, coloridas=0, inseridas=0, freq=0, reducao=0, exclusoes=0,
             alteracao=0, horario=0, reducao_horario=0)

    # ---- 1) Mapeia alterações da TimeTable (pela CHAVE) ----
    dic_cor, dic_count, dic_tipo, dic_linha_tt = {}, {}, {}, {}
    for i in range(lin_cab_tt + 1, ult_lin_tt + 1):
        if i not in tt_dados:
            continue
        tipo_alt = texto_bruto(tt_v(i, cols_tt[6]))
        if tipo_alt:
            chave = norm_chave(tt_v(i, col_chave_tt))
            if chave and chave not in dic_cor:
                dic_cor[chave] = tt_cor(i, cols_tt[6])
                dic_count[chave] = 0
                dic_tipo[chave] = tipo_alt
                dic_linha_tt[chave] = i
                q["com_tipo"] += 1

    # ---- 2) Mapeia linhas EXISTENTES do Cenário (pela CHAVE) ----
    dic_linhas_cen = {}
    for i in range(lin_cab_cen + 1, ult_lin_cen + 1):
        chave = norm_chave(cen_v(i, col_chave_cen))
        if not chave:
            continue
        dic_linhas_cen.setdefault(chave, []).append(i)

        if chave in dic_cor:
            tipo_alt = dic_tipo[chave]
            if eh_exclusao(tipo_alt):
                lote.pintar(i, 1, ult_col_cen, COR["VERMELHO_ESCURO"])
                if eh_exclusao_servico(tipo_alt):
                    lote.pintar(i, cols_cen[1], cols_cen[1], COR["VINHO_SERVICO"])
                lote.fonte(i, 1, ult_col_cen, "#FFFFFF")
                lote.centralizar(i, 1, ult_col_cen)
                q["exclusoes"] += 1
                dic_count[chave] += 1
            elif not (eh_inclusao_frequencia(tipo_alt) or eh_reducao_frequencia(tipo_alt) or
                      eh_alteracao_frequencia(tipo_alt) or eh_alteracao_horario(tipo_alt) or
                      eh_reducao_freq_e_alt_horario(tipo_alt)):
                lote.pintar(i, 1, ult_col_cen, dic_cor[chave])
                q["coloridas"] += 1
                dic_count[chave] += 1

    proxima_linha_cen = ult_lin_cen + 1
    insercoes = []

    # ---- 2c) INCLUSÃO DE FREQUÊNCIA (planejamento) ----
    for vk, tipo in dic_tipo.items():
        if not eh_inclusao_frequencia(tipo):
            continue
        linha_orig = dic_linha_tt[vk]

        if vk in dic_linhas_cen:
            # Divide as linhas da chave em blocos (quando o rank do dia deixa de crescer)
            blocks, curr, prev = [], [], 0
            for r in dic_linhas_cen[vk]:
                cur = rank_linha(r)
                if cur <= prev and curr:
                    blocks.append(curr)
                    curr = []
                curr.append(r)
                prev = cur
            if curr:
                blocks.append(curr)

            for dia_tt in range(12, 19):
                if texto_bruto(tt_v(linha_orig, cols_tt[dia_tt])).upper() == "S":
                    r_new = day_rank[dia_tt - 12]
                    col_dia = cols_cen[dia_tt - 12 + 9]
                    for blk in blocks:
                        if any(eh_marcacao_s(cen_v(r_in, col_dia)) for r_in in blk):
                            continue
                        alvo = None
                        for r_in in blk:
                            if rank_linha(r_in) > r_new:
                                alvo = r_in
                                break
                        if alvo is None:
                            alvo = blk[-1] + 1
                        insercoes.append(dict(targetRow=alvo, dayRank=r_new,
                                              diaIdxTT=dia_tt, linhaOrig=linha_orig))
            dic_count[vk] = 1
        else:
            for dia_tt in range(12, 19):
                if texto_bruto(tt_v(linha_orig, cols_tt[dia_tt])).upper() == "S":
                    insercoes.append(dict(targetRow=proxima_linha_cen,
                                          dayRank=day_rank[dia_tt - 12],
                                          diaIdxTT=dia_tt, linhaOrig=linha_orig))
            dic_count[vk] = 1

    # ---- 2d) REDUÇÃO DE FREQUÊNCIA ----
    for vk, tipo in dic_tipo.items():
        if not eh_reducao_frequencia(tipo) or vk not in dic_linhas_cen:
            continue
        linha_orig = dic_linha_tt[vk]
        for dia_tt in range(12, 19):
            v_tt, b_tt = tt_v(linha_orig, cols_tt[dia_tt]), tt_cor(linha_orig, cols_tt[dia_tt])
            if eh_marcacao_reducao(v_tt, b_tt):
                col_c = cols_cen[dia_tt - 12 + 9]
                for r in dic_linhas_cen[vk]:
                    if eh_marcacao_s(cen_v(r, col_c)):
                        set_cen(r, col_c, "X")
                        lote.pintar(r, 1, ult_col_cen, COR["VERMELHO_CLARO"])
                        lote.centralizar(r, 1, ult_col_cen)
                        lote.pintar(r, col_c, col_c, b_tt)
                        q["reducao"] += 1
        dic_count[vk] = 1

    # ---- 2d-bis) ALTERAÇÃO DE FREQUÊNCIA ----
    linhas_alt = set()
    for vk, tipo in dic_tipo.items():
        if not eh_alteracao_frequencia(tipo) or vk not in dic_linhas_cen:
            continue
        linha_orig = dic_linha_tt[vk]
        dias_saem, dias_entram = [], []
        for dia_tt in range(12, 19):
            v_tt, b_tt = tt_v(linha_orig, cols_tt[dia_tt]), tt_cor(linha_orig, cols_tt[dia_tt])
            if eh_marcacao_reducao(v_tt, b_tt):
                dias_saem.append(dia_tt)
            elif texto_bruto(v_tt).upper() == "S" and cor_eh_verde(b_tt):
                dias_entram.append(dia_tt)

        for dia_sai in dias_saem:
            col_sai = cols_cen[dia_sai - 12 + 9]
            b_sai = tt_cor(linha_orig, cols_tt[dia_sai])
            for r in dic_linhas_cen[vk]:
                if eh_marcacao_s(cen_v(r, col_sai)):
                    if r not in linhas_alt:
                        linhas_alt.add(r)
                        lote.pintar(r, 1, ult_col_cen, COR["LARANJA"])
                    set_cen(r, col_sai, "X")
                    lote.pintar(r, col_sai, col_sai, b_sai)
                    lote.centralizar(r, col_sai, col_sai)
                    for dia_entra in dias_entram:
                        col_entra = cols_cen[dia_entra - 12 + 9]
                        set_cen(r, col_entra, "S")
                        lote.pintar(r, col_entra, col_entra, tt_cor(linha_orig, cols_tt[dia_entra]))
                        lote.centralizar(r, col_entra, col_entra)
                    q["alteracao"] += 1
        dic_count[vk] = 1

    # ---- 2d-ter) ALTERAÇÃO DE HORÁRIO ----
    for vk, tipo in dic_tipo.items():
        if not eh_alteracao_horario(tipo) or vk not in dic_linhas_cen:
            continue
        h = ler_alteracao_horario(dic_linha_tt[vk], cols_tt, tt_v, tt_cor)
        if h["altPartida"] or h["altChegada"]:
            for r in dic_linhas_cen[vk]:
                lote.pintar(r, 1, ult_col_cen, COR["AMARELO_CLARO"])
                aplicar_horario(lote, r, cols_cen, h)
                q["horario"] += 1
        dic_count[vk] = 1

    # ---- 2d-quater) REDUÇÃO DE FREQUÊNCIA E ALTERAÇÃO DE HORÁRIO ----
    for vk, tipo in dic_tipo.items():
        if not eh_reducao_freq_e_alt_horario(tipo) or vk not in dic_linhas_cen:
            continue
        linha_orig = dic_linha_tt[vk]
        dia_sai = {d: eh_marcacao_reducao(tt_v(linha_orig, cols_tt[d]), tt_cor(linha_orig, cols_tt[d]))
                   for d in range(12, 19)}
        h = ler_alteracao_horario(linha_orig, cols_tt, tt_v, tt_cor)

        for r in dic_linhas_cen[vk]:
            alterada = h["altPartida"] or h["altChegada"]
            if not alterada:
                alterada = any(dia_sai[d] and eh_marcacao_s(cen_v(r, cols_cen[d - 12 + 9]))
                               for d in range(12, 19))
            if alterada:
                lote.pintar(r, 1, ult_col_cen, COR["AZUL_CLARO"])
                for d in range(12, 19):
                    if dia_sai[d]:
                        col_c = cols_cen[d - 12 + 9]
                        if eh_marcacao_s(cen_v(r, col_c)):
                            b_tt = tt_cor(linha_orig, cols_tt[d])
                            set_cen(r, col_c, "X")
                            lote.pintar(r, col_c, col_c,
                                        b_tt if cor_eh_vermelho(b_tt) else COR["VERMELHO"])
                            lote.centralizar(r, col_c, col_c)
                aplicar_horario(lote, r, cols_cen, h)
                q["reducao_horario"] += 1
        dic_count[vk] = 1

    # ---- Grava o que mexeu em linhas existentes (antes de inserir linhas) ----
    lote.aplicar(ws_cen)
    lote.limpar()

    # ---- Inserções de frequência ----
    n_ins = len(insercoes)
    if n_ins > 0:
        insercoes.sort(key=lambda it: (-it["targetRow"], -it["dayRank"]))
        inserir_linhas(ws_cen, insercoes)

        for idx, it in enumerate(insercoes):
            r_final = it["targetRow"] + (n_ins - idx - 1)
            dia_cen_idx = it["diaIdxTT"] - 12 + 9
            escrever_linha_nova(lote, r_final, tt_v, it["linhaOrig"], cols_tt, cols_cen,
                                dia_cen_idx, COR["VERDE_CLARO"], ult_col_cen, True)
            lote.valor(r_final, col_chave_cen, chave_nova(r_final, it["linhaOrig"]))
            cor_dia = tt_cor(it["linhaOrig"], cols_tt[it["diaIdxTT"]])
            if not cor_eh_branca(cor_dia):
                lote.pintar(r_final, cols_cen[dia_cen_idx], cols_cen[dia_cen_idx], cor_dia)
            q["freq"] += 1
            proxima_linha_cen += 1

    # ---- 2e) INCLUSÃO padrão (respeita a ordem física das colunas de dias) ----
    for vk, cor in dic_cor.items():
        if dic_count[vk] == 0 and eh_inclusao(dic_tipo[vk]):
            linha_orig = dic_linha_tt[vk]
            inc_serv = eh_inclusao_servico(dic_tipo[vk])
            cor_linha = COR["VERDE_INCLUSAO_SERVICO"] if inc_serv else cor

            for rank_loop in range(1, 8):
                for dia_rel in range(7):
                    if day_rank[dia_rel] == rank_loop:
                        dia_tt = 12 + dia_rel
                        if texto_bruto(tt_v(linha_orig, cols_tt[dia_tt])).upper() == "S":
                            escrever_linha_nova(lote, proxima_linha_cen, tt_v, linha_orig, cols_tt,
                                                cols_cen, dia_rel + 9, cor_linha, ult_col_cen, False)
                            lote.valor(proxima_linha_cen, col_chave_cen,
                                       chave_nova(proxima_linha_cen, linha_orig))
                            if inc_serv:
                                lote.pintar(proxima_linha_cen, cols_cen[1], cols_cen[1],
                                            COR["VERDE_SERVICO_CELULA"])
                            q["inseridas"] += 1
                            proxima_linha_cen += 1
                        break
            dic_count[vk] = 1

    lote.aplicar(ws_cen)
    wb.save(saida)

    # ---- 3) Relatório ----
    nao_enc = [vk for vk, qtd in dic_count.items() if qtd == 0]
    linhas = [
        "Linhas com 'Tipo de Alteração' na TimeTable: %d" % q["com_tipo"],
        "Linhas coloridas em '%s': %d" % (nome_aba_cenario, q["coloridas"]),
        "Exclusões (linha mantida em vermelho escuro): %d" % q["exclusoes"],
        "Linhas inseridas (Inclusão): %d" % q["inseridas"],
        "Linhas adicionadas (Inclusão de Frequência): %d" % q["freq"],
        "Linhas alteradas (Redução de Frequência): %d" % q["reducao"],
        "Linhas alteradas (Alteração de Frequência): %d" % q["alteracao"],
        "Linhas alteradas (Alteração de Horário): %d" % q["horario"],
        "Linhas alteradas (Redução de Frequência e Alteração de Horário): %d" % q["reducao_horario"],
        "Chaves não encontradas no Cenário: %d" % len(nao_enc),
    ]
    linhas += ["    - " + vk for vk in nao_enc[:15]]
    if len(nao_enc) > 15:
        linhas.append("    ...")
    print("\n".join(linhas))
    print("\nArquivo gerado: %s" % saida)


# ============================ HORÁRIOS / LINHAS NOVAS ============================

def ler_alteracao_horario(linha_orig, cols_tt, tt_v, tt_cor):
    v_part, v_cheg = tt_v(linha_orig, cols_tt[9]), tt_v(linha_orig, cols_tt[10])
    alt_part = cor_eh_amarelo(tt_cor(linha_orig, cols_tt[9]))
    alt_cheg = cor_eh_amarelo(tt_cor(linha_orig, cols_tt[10]))

    # Se nenhuma estiver amarela, considera a(s) que tiver(em) valor preenchido
    if not alt_part and not alt_cheg:
        alt_part = len(texto_bruto(v_part)) > 0
        alt_cheg = len(texto_bruto(v_cheg)) > 0
    # Nunca apaga um horário do Cenário com uma célula vazia da TimeTable
    if not texto_bruto(v_part):
        alt_part = False
    if not texto_bruto(v_cheg):
        alt_cheg = False
    return dict(altPartida=alt_part, altChegada=alt_cheg, vPart=v_part, vCheg=v_cheg)


def aplicar_horario(lote, r, cols_cen, h):
    for flag, chave, col in (("altPartida", "vPart", cols_cen[6]), ("altChegada", "vCheg", cols_cen[7])):
        if h[flag]:
            v = valor_ou_branco(h[chave])
            lote.valor(r, col, v)
            if texto_bruto(v):
                lote.formato_hora(r, col)
            lote.pintar(r, col, col, COR["AMARELO"])


def escrever_linha_nova(lote, r, tt_v, l_orig, cols_tt, cols_cen, dia_cen_idx, cor,
                        ult_col_cen, resetar_fonte):
    """Linha nova do Cenário com 1 único "S" no dia indicado."""
    v = lambda k: tt_v(l_orig, cols_tt[k])

    lote.valor(r, cols_cen[0], valor_ou_branco(combinar_exibicao(v(0), v(1))))
    lote.valor(r, cols_cen[1], valor_ou_branco(combinar_exibicao(v(2), v(3))))
    lote.valor(r, cols_cen[2], valor_ou_branco(v(4)))
    lote.valor(r, cols_cen[3], valor_ou_branco(v(5)))
    lote.valor(r, cols_cen[4], valor_ou_branco(v(7)))
    lote.valor(r, cols_cen[5], valor_ou_branco(v(8)))

    for idx_cen, idx_tt in ((6, 9), (7, 10)):
        h = valor_ou_branco(v(idx_tt))
        lote.valor(r, cols_cen[idx_cen], h)
        if texto_bruto(h):
            lote.formato_hora(r, cols_cen[idx_cen])

    f = texto_bruto(valor_ou_branco(v(11)))
    lote.valor(r, cols_cen[8], "" if not f else (f if f[:2].upper() == "DD" else "DD" + f))

    for k_dia in range(7):
        lote.valor(r, cols_cen[9 + k_dia], "")
    lote.valor(r, cols_cen[dia_cen_idx], "S")

    lote.pintar(r, 1, ult_col_cen, cor)
    lote.centralizar(r, 1, ult_col_cen)
    if resetar_fonte:
        lote.fonte(r, 1, ult_col_cen, "#000000")


# ============================ LINHA DE COMANDO ============================

def main():
    ap = argparse.ArgumentParser(description="Aplica as alterações da TimeTable na aba de cenário.")
    ap.add_argument("arquivo", help="arquivo .xlsx de entrada")
    ap.add_argument("cenario", help="nome da aba de cenário (ex.: 'Cenário 1')")
    ap.add_argument("-o", "--saida", help="arquivo de saída (padrão: <arquivo>_processado.xlsx)")
    ap.add_argument("--timetable", default=ABA_TIMETABLE, help="nome da aba TimeTable")
    a = ap.parse_args()

    entrada = Path(a.arquivo)
    saida = Path(a.saida) if a.saida else entrada.with_name(entrada.stem + "_processado" + entrada.suffix)
    if saida.resolve() == entrada.resolve():
        sys.exit("O arquivo de saída não pode ser o mesmo da entrada.")
    try:
        executar(entrada, saida, a.cenario, a.timetable)
    except ErroProcesso as e:
        sys.exit("Atenção: %s" % e)