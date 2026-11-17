
import sys
import time
import traceback
import copy
import unicodedata

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import PatternFill

from PySide6.QtCore import (
    QAbstractTableModel,
    QModelIndex,
    QObject,
    Qt,
    QThread,
    Signal,
    Slot,
)

from PySide6.QtGui import QColor

from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHeaderView,
    QLabel,
    QMessageBox,
    QProgressDialog,
    QPushButton,
    QTableView,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from excel_io import (
    carregar_cenario,
    carregar_timetable,
    salvar_cenario,
)

from regras_cenario import _processar as processar_em_memoria


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

    def __init__(self, df=None, bgs=None, fontes=None):
        super().__init__()

        self.definir(
            df,
            bgs,
            fontes
        )

    def definir(self, df, bgs=None, fontes=None):

        self.beginResetModel()

        self.df = df
        self.bgs = bgs or []
        self.fontes = fontes or []

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

        if role != Qt.DisplayRole or self.df is None:
            return None

        if orientation == Qt.Horizontal:
            return str(self.df.columns[section])

        return str(section + 1)

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
# ABA IMPORTADOR
# ============================================================

class AbaImportador(QWidget):
    """Importa Excel: o cabeçalho é localizado automaticamente."""

    importado = Signal()

    def __init__(self, texto_botao: str, tipo: str):

        super().__init__()

        self.tipo = tipo

        self.caminho_arquivo = None
        self.planilha_encontrada = None
        self.header_row = 2

        self.df = None
        self.bgs = None
        self.fontes = None

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

        self.label_arquivo = QLabel(
            "Nenhum arquivo selecionado"
        )

        self.label_arquivo.setWordWrap(True)

        self.modelo = PlanilhaModel()

        self.tabela = QTableView()

        self.tabela.setModel(
            self.modelo
        )

        self.tabela.setEditTriggers(
            QTableView.NoEditTriggers
        )

        self.tabela.setAlternatingRowColors(False)

        self.tabela.verticalHeader().setDefaultSectionSize(22)

        header = self.tabela.horizontalHeader()

        header.setSectionResizeMode(
            QHeaderView.Interactive
        )

        header.setStretchLastSection(False)

        layout = QVBoxLayout(self)

        layout.addWidget(self.botao_importar)
        layout.addWidget(self.label_arquivo)
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
            self.fontes
        )

        self.tabela.resizeColumnsToContents()

        self.importado.emit()

    def atualizar(self, df, bgs, fontes):
        """Mostra o resultado processado, sem alterar os dados originais."""

        self.df = df
        self.bgs = bgs
        self.fontes = fontes

        self.modelo.definir(
            df,
            bgs,
            fontes
        )

        self.tabela.resizeColumnsToContents()

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
            "Importador de Excel - TimeTable e Cenário"
        )

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

        dlg = QProgressDialog(
            "Iniciando…",
            "",
            0,
            100,
            self
        )

        dlg.setCancelButton(None)

        dlg.setWindowTitle(
            "Processando cenário"
        )

        dlg.setWindowModality(
            Qt.WindowModal
        )

        dlg.setMinimumDuration(0)
        dlg.setAutoClose(False)
        dlg.setAutoReset(False)
        dlg.setValue(0)

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

if __name__ == "__main__":

    app = QApplication(sys.argv)

    janela = JanelaPrincipal()

    janela.show()

    sys.exit(app.exec())