import sys
import traceback
import tempfile
import contextlib
import io
import copy
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
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
    QPushButton,
    QTableView,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from excel_io import carregar_cenario, carregar_timetable, salvar_cenario
from regras_cenario import executar as executar_regras_cenario


@dataclass
class ResultadoProcessamento:
    df: object
    bgs: list
    fontes: list
    relatorio: str


def _valor_excel(valor):
    """Converte valores do pandas para tipos aceitos pelo openpyxl."""
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


def _montar_aba(ws, df, bgs=None, fontes=None):
    """Escreve DataFrame e cores em uma aba com cabeçalho na primeira linha."""
    if df is None:
        raise ValueError("Uma das planilhas não foi importada corretamente.")
    for c, nome in enumerate(df.columns, 1):
        ws.cell(1, c, str(nome))
    for r in range(len(df)):
        for c in range(df.shape[1]):
            cell = ws.cell(r + 2, c + 1, _valor_excel(df.iat[r, c]))
            if bgs and r < len(bgs) and c < len(bgs[r]):
                cor = _cor_hex(bgs[r][c])
                cell.fill = PatternFill(fill_type="solid", fgColor="FF" + cor[1:])
            if fontes and r < len(fontes) and c < len(fontes[r]) and fontes[r][c]:
                fonte = copy.copy(cell.font)
                fonte.color = _cor_hex(fontes[r][c], padrao="#000000")
                cell.font = fonte


def _ler_aba_processada(ws):
    """Lê valores, preenchimentos e cores de fonte do resultado das regras."""
    cabecalho = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    linhas, bgs, fontes = [], [], []
    for r in range(2, ws.max_row + 1):
        valores = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        # Ignora linhas totalmente vazias, sem descartar linhas que tenham apenas marcação em dias.
        if not any(v is not None for v in valores):
            continue
        linhas.append(valores)
        cores_linha, fontes_linha = [], []
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            fill = cell.fill
            if fill and fill.fill_type == "solid" and fill.fgColor.type == "rgb":
                rgb = str(fill.fgColor.rgb)[-6:]
                cores_linha.append("#" + rgb)
            else:
                cores_linha.append("#FFFFFF")
            cor_fonte = cell.font.color
            if cor_fonte is not None and cor_fonte.type == "rgb":
                fontes_linha.append("#" + str(cor_fonte.rgb)[-6:])
            else:
                fontes_linha.append(None)
        bgs.append(cores_linha)
        fontes.append(fontes_linha)
    return pd.DataFrame(linhas, columns=cabecalho), bgs, fontes


def processar_cenario(df_tt, bg_tt, df_cen, bg_cen, cols_tt=None, cols_cen=None, fontes_cen=None):
    """Adapta os dados importados pela interface ao processador oficial de regras.

    O módulo regras_cenario trabalha com arquivos Excel. Por isso, os DataFrames
    importados pela interface são montados em um arquivo temporário, processados
    por executar(), e o resultado é devolvido à interface como DataFrame e cores.
    """
    with tempfile.TemporaryDirectory(prefix="processamento_cenario_") as pasta:
        entrada = Path(pasta) / "entrada.xlsx"
        saida = Path(pasta) / "saida.xlsx"

        wb = Workbook()
        ws_tt = wb.active
        ws_tt.title = "TimeTable"
        _montar_aba(ws_tt, df_tt, bg_tt)
        ws_cen = wb.create_sheet("Cenario")
        _montar_aba(ws_cen, df_cen, bg_cen, fontes_cen)
        wb.save(entrada)
        wb.close()

        relatorio = io.StringIO()
        with contextlib.redirect_stdout(relatorio):
            executar_regras_cenario(entrada, saida, "Cenario", "TimeTable")

        wb_resultado = load_workbook(saida, data_only=True)
        try:
            df_resultado, bgs_resultado, fontes_resultado = _ler_aba_processada(
                wb_resultado["Cenario"]
            )
        finally:
            wb_resultado.close()

        return ResultadoProcessamento(
            df=df_resultado,
            bgs=bgs_resultado,
            fontes=fontes_resultado,
            relatorio=relatorio.getvalue().strip(),
        )


class PlanilhaModel(QAbstractTableModel):
    """Modelo virtual: a tabela só pinta as células visíveis."""

    def __init__(self, df=None, bgs=None, fontes=None):
        super().__init__()
        self.definir(df, bgs, fontes)

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

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole or self.df is None:
            return None
        if orientation == Qt.Horizontal:
            return str(self.df.columns[section])
        return str(section + 1)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or self.df is None:
            return None
        r, c = index.row(), index.column()
        if role == Qt.DisplayRole:
            valor = self.df.iat[r, c]
            if valor is None:
                return ""
            try:
                if valor != valor:  # NaN
                    return ""
            except Exception:
                pass
            return str(valor)
        if role == Qt.BackgroundRole:
            if r < len(self.bgs) and c < len(self.bgs[r]):
                return QColor(self.bgs[r][c])
        if role == Qt.ForegroundRole:
            if r < len(self.fontes) and c < len(self.fontes[r]) and self.fontes[r][c]:
                return QColor(self.fontes[r][c])
            # O fundo é sempre pintado explicitamente (branco por padrão);
            # sem isto, em tema escuro o texto ficaria branco sobre branco.
            return QColor("#000000")
        if role == Qt.TextAlignmentRole:
            return Qt.AlignCenter
        return None


class Worker(QObject):
    concluido = Signal(object)
    falhou = Signal(str)

    def __init__(self, fn):
        super().__init__()
        self.fn = fn

    @Slot()
    def run(self):
        try:
            resultado = self.fn()
        except Exception as exc:
            traceback.print_exc()
            self.falhou.emit(f"{type(exc).__name__}: {exc}")
            return

        self.concluido.emit(resultado)


class Execucao(QObject):
    """
    Executa `fn` em uma QThread e chama on_ok / on_err NA THREAD DA INTERFACE.

    Os sinais do Worker são ligados a slots (métodos) deste QObject, que vive
    na thread principal. Ligar o sinal a uma função/closure solta fazia o
    callback rodar na thread de trabalho, e mexer em widgets (QMessageBox,
    setEnabled, setText...) fora da thread principal é comportamento
    indefinido no Qt.
    """

    _ativas: set = set()

    def __init__(self, fn, on_ok=None, on_err=None):
        super().__init__()
        self._on_ok = on_ok
        self._on_err = on_err
        self._thread = QThread()
        self._worker = Worker(fn)
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.concluido.connect(self._receber_ok)
        self._worker.falhou.connect(self._receber_erro)
        self._thread.finished.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._finalizar)

        Execucao._ativas.add(self)  # mantém a referência viva até terminar

    def iniciar(self):
        self._thread.start()

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


class AbaImportador(QWidget):
    """Importa Excel: cabeçalho na linha 1 (ou a localizada) e dados a partir da seguinte."""

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
        # Cópia do que foi importado: "Aplicar alterações" sempre parte daqui.
        # Sem isso, aplicar duas vezes reprocessava o cenário já processado
        # e duplicava as linhas inseridas.
        self.df_original = None
        self.bgs_original = None
        self.fontes_original = None
        self.cols = None

        self.botao_importar = QPushButton(texto_botao)
        self.botao_importar.clicked.connect(self.escolher_arquivo)
        self.label_arquivo = QLabel("Nenhum arquivo selecionado")
        self.label_arquivo.setWordWrap(True)

        self.modelo = PlanilhaModel()
        self.tabela = QTableView()
        self.tabela.setModel(self.modelo)
        self.tabela.setEditTriggers(QTableView.NoEditTriggers)
        self.tabela.setAlternatingRowColors(False)
        self.tabela.verticalHeader().setDefaultSectionSize(22)
        header = self.tabela.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setStretchLastSection(False)

        layout = QVBoxLayout(self)
        layout.addWidget(self.botao_importar)
        layout.addWidget(self.label_arquivo)
        layout.addWidget(self.tabela)

    def escolher_arquivo(self):
        # .xls não é suportado pelo openpyxl, por isso não é oferecido.
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
        self.label_arquivo.setText(f"Lendo {Path(caminho).name}…")

        def tarefa():
            if self.tipo == "tt":
                dados = carregar_timetable(caminho)
            elif self.tipo == "cen":
                dados = carregar_cenario(caminho)
            else:
                raise ValueError(f"Tipo de planilha inválido: {self.tipo}")
            dados["caminho"] = caminho
            return dados

            # Neste modelo, a primeira linha é informativa e os cabeçalhos
            # das abas ficam na segunda linha. A cópia temporária remove
            # somente essa primeira linha, preservando o arquivo original.
            sufixo = Path(caminho).suffix.lower()
            with tempfile.TemporaryDirectory(prefix="leitura_planilha_") as pasta:
                copia = Path(pasta) / f"arquivo{Path(caminho).suffix}"
                wb = load_workbook(
                    caminho, data_only=False, keep_vba=(sufixo == ".xlsm")
                )
                try:
                    for ws in wb.worksheets:
                        if ws.max_row >= 1:
                            ws.delete_rows(1, 1)
                    wb.save(copia)
                finally:
                    wb.close()

                if self.tipo == "tt":
                    dados = carregar_timetable(str(copia))
                else:
                    dados = carregar_cenario(str(copia))

                # Mantém o caminho original para a rotina de salvamento.
                dados["caminho"] = caminho
                return dados

        self._rodar(tarefa, self._aplicar_carga, self._falha_carga)

    def _falha_carga(self, msg: str):
        self.label_arquivo.setText("Falha na importação.")
        QMessageBox.critical(self, "Erro", f"Não foi possível abrir o arquivo:\n{msg}")

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
        n = 0 if self.df is None else len(self.df)
        self.label_arquivo.setText(
            f"{Path(self.caminho_arquivo).name}  |  aba {self.planilha_encontrada}  |  "
            f"{n} linhas (cabeçalho na linha {self.header_row})"
        )
        self.modelo.definir(self.df, self.bgs, self.fontes)
        self.tabela.resizeColumnsToContents()
        self.importado.emit()

    def atualizar(self, df, bgs, fontes):
        """Mostra o resultado processado, sem alterar os dados originais."""
        self.df = df
        self.bgs = bgs
        self.fontes = fontes
        self.modelo.definir(df, bgs, fontes)

    def _rodar(self, fn, on_ok, on_err):
        QApplication.setOverrideCursor(Qt.WaitCursor)
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

        Execucao(fn, ok, err).iniciar()


class JanelaPrincipal(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Importador de Excel - TimeTable e Cenário")
        self.resize(1100, 700)
        self.resultado: ResultadoProcessamento | None = None

        self.aba_timetable = AbaImportador("Importar TimeTable", "tt")
        self.aba_cenario = AbaImportador("Importar Cenário", "cen")
        # Nova importação invalida o resultado anterior.
        self.aba_timetable.importado.connect(self._invalidar_resultado)
        self.aba_cenario.importado.connect(self._invalidar_resultado)

        self.botao_aplicar = QPushButton("Aplicar alterações")
        self.botao_aplicar.clicked.connect(self.aplicar_alteracoes)
        self.botao_baixar = QPushButton("Baixar cenário processado")
        self.botao_baixar.setEnabled(False)
        self.botao_baixar.clicked.connect(self.baixar_cenario)
        self.label_status = QLabel("")
        self.label_status.setWordWrap(True)

        self.abas = QTabWidget()
        self.abas.addTab(self.aba_timetable, "TimeTable")
        self.abas.addTab(self.aba_cenario, "Cenário")

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

    def _rodar(self, fn, on_ok, on_err=None, botao_bloquear=None):
        if botao_bloquear is None:
            botao_bloquear = self.botao_aplicar

        QApplication.setOverrideCursor(Qt.WaitCursor)
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
                QMessageBox.critical(self, "Erro", msg)

        Execucao(fn, ok, err).iniciar()

    def aplicar_alteracoes(self):
        tt = self.aba_timetable
        cen = self.aba_cenario
        if tt.df_original is None:
            QMessageBox.warning(self, "TimeTable não importado", "Importe o TimeTable antes de aplicar as alterações.")
            return
        if cen.df_original is None:
            QMessageBox.warning(self, "Cenário não importado", "Importe o Cenário antes de aplicar as alterações.")
            return

        self.label_status.setText("Processando cenário…")

        # Sempre a partir dos dados ORIGINAIS importados.
        df_tt, bg_tt = tt.df_original, tt.bgs_original
        df_cen, bg_cen, fontes_cen = cen.df_original, cen.bgs_original, cen.fontes_original
        cols_tt, cols_cen = tt.cols, cen.cols

        def tarefa():
            return processar_cenario(
                df_tt, bg_tt, df_cen, bg_cen, cols_tt, cols_cen, fontes_cen,
            )

        def erro(mensagem):
            self.label_status.setText("Falha ao processar o cenário.")
            QMessageBox.critical(self, "Erro ao processar", mensagem)

        self._rodar(tarefa, self._on_processado, erro)

    def _on_processado(self, resultado: ResultadoProcessamento):
        self.resultado = resultado
        self.aba_cenario.atualizar(resultado.df, resultado.bgs, resultado.fontes)
        self.abas.setCurrentWidget(self.aba_cenario)
        self.botao_baixar.setEnabled(True)
        self.label_status.setText(resultado.relatorio.replace("\n", "  |  "))

    def baixar_cenario(self):
        if self.resultado is None:
            QMessageBox.warning(
                self,
                "Atenção",
                "Não existe cenário processado para salvar.",
            )
            return

        origem = Path(self.aba_cenario.caminho_arquivo or "Cenario.xlsx")
        sugerido = str(origem.with_name(origem.stem + "_processado.xlsx"))

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

        destino = Path(caminho).expanduser().resolve()
        resultado = self.resultado

        kwargs = {
            "caminho_origem": self.aba_cenario.caminho_arquivo,
            "nome_aba": self.aba_cenario.planilha_encontrada or "Cenario",
            "header_row": self.aba_cenario.header_row or 1,
            "df": resultado.df,
            "bgs": resultado.bgs,
            "fontes": resultado.fontes,
            "caminho_saida": str(destino),
        }

        self.label_status.setText("Salvando arquivo...")

        def tarefa():
            salvar_cenario(**kwargs)

            if not destino.is_file():
                raise FileNotFoundError(
                    f"O arquivo não foi encontrado após a gravação: {destino}"
                )

            if destino.stat().st_size == 0:
                raise IOError("O arquivo salvo está vazio.")

            return str(destino)

        def sucesso(saida):
            self.label_status.setText(f"Arquivo salvo: {saida}")
            # Antes havia "\\n" (barra literal) e a mensagem não quebrava linha.
            QMessageBox.information(
                self,
                "Salvamento concluído",
                f"O arquivo foi salvo com sucesso em:\n{saida}",
            )

        def erro(mensagem):
            self.label_status.setText("Falha ao salvar o arquivo.")
            QMessageBox.critical(
                self,
                "Erro ao salvar",
                f"Não foi possível salvar o cenário:\n\n{mensagem}",
            )

        self._rodar(tarefa, sucesso, erro, botao_bloquear=self.botao_baixar)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    janela = JanelaPrincipal()
    janela.show()
    sys.exit(app.exec())
