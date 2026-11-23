#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Personalização visual da aplicação PySide6.

Integração no main.py:
    from estilo_app import (
        aplicar_estilo,
        aplicar_fundo_blur,
        DialogoProgressoLinha,
        QMessageBox,
        QInputDialog,
    )

    app = QApplication(sys.argv)
    aplicar_estilo(app)

    janela = JanelaPrincipal()
    aplicar_fundo_blur(janela)
    janela.show()

O fundo usa gradientes radiais animados com transições suaves para criar um
efeito de blur em movimento nas cores coral e verde-água. Cada mancha percorre
a tela seguindo uma curva de Lissajous.

As tabelas mantêm o delegate padrão do Qt para preservar as cores de fundo e
fonte fornecidas pelo PlanilhaModel a partir do Excel.

Pop-ups: QMessageBox e QInputDialog exportados por este módulo têm as mesmas
chamadas estáticas das classes do Qt (critical, warning, information, getText,
getItem), mas aparecem em branco translúcido. DialogoProgressoLinha substitui
o QProgressDialog, com barra de progresso em linha fina.
"""

import math

from PySide6.QtCore import QEvent, QPointF, Qt, QTimer
from PySide6.QtGui import QFont, QPalette, QColor, QPainter, QRadialGradient
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QLabel,
    QLineEdit,
    QProgressBar,
    QPushButton,
    QTableView,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtWidgets import (
    QInputDialog as _QInputDialog,
    QMessageBox as _QMessageBox,
)


# ============================================================
# PALETA DE CORES
# ============================================================

CORES = {
    "primaria": "#0E3862",
    "primaria_hover": "#174B7A",
    "primaria_pressed": "#092947",
    "coral": "#F16664",
    "verde_agua": "#76B1B2",
    "fundo": "#F3F6F9",
    "superficie": "#FFFFFF",
    "superficie_secundaria": "#EAF0F5",
    "texto": "#263445",
    "texto_secundario": "#607387",
    "borda": "#D9E1E8",
    "selecao": "#DCEBEC",
    "selecao_texto": "#0E3862",
    "desabilitado": "#AAB7C4",
    "desabilitado_texto": "#EDF1F4",
    "alternada": "#F7F9FB",
    "branco": "#FFFFFF",
}

FONTE_FAMILIA = "Segoe UI"
FONTE_TAMANHO = 10

# Transparência branca da aba selecionada (0 = invisível; 255 = opaca).
OPACIDADE_ABA_SELECIONADA = 145

# Transparência branca dos pop-ups (0 = invisível; 255 = opaca).
OPACIDADE_POPUP = 200            # mensagens e entradas de texto
OPACIDADE_POPUP_PROGRESSO = 170  # diálogo de progresso

# Intensidade das manchas de cor do fundo (0 a 255).
INTENSIDADE_BLUR = 115

# Tamanho das manchas em relação ao menor lado da janela (1.0 = tamanho base).
TAMANHO_BLUR = 1.25

# Animação do fundo. Defina ANIMACAO_ATIVA = False para deixar as manchas paradas.
ANIMACAO_ATIVA = True
INTERVALO_ANIMACAO_MS = 33          # aproximadamente 30 atualizações por segundo
VELOCIDADE_ANIMACAO = 0.01          # maior = mais rápido
AMPLITUDE_X = 0.50                  # deslocamento horizontal relativo
AMPLITUDE_Y = 0.50                  # deslocamento vertical relativo


# ============================================================
# FUNDO COM GRADIENTES RADIAIS (EFEITO BLUR)
# ============================================================

class FundoBlur(QWidget):
    """Camada de fundo com manchas coloridas e bordas muito suaves."""

    def __init__(self, janela: QWidget):
        super().__init__(janela)

        self.setObjectName("fundoBlur")
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setGeometry(janela.rect())

        # A camada acompanha o tamanho da janela principal.
        janela.installEventFilter(self)

        # O timer anima somente esta camada, sem bloquear a interface.
        self._fase = 0.0
        self._timer = QTimer(self)
        self._timer.setInterval(INTERVALO_ANIMACAO_MS)
        self._timer.timeout.connect(self._avancar_animacao)
        if ANIMACAO_ATIVA and janela.isVisible():
            self._timer.start()

        # Mantém o fundo atrás dos demais componentes da interface.
        self.lower()

    def _avancar_animacao(self):
        """Avança a fase e solicita apenas o redesenho do fundo."""
        self._fase += VELOCIDADE_ANIMACAO
        self.update()

    def eventFilter(self, watched, event):
        if watched is self.parentWidget() and event.type() == QEvent.Resize:
            self.setGeometry(self.parentWidget().rect())
        elif watched is self.parentWidget() and event.type() == QEvent.Show:
            self.setGeometry(self.parentWidget().rect())
            self.lower()
            if ANIMACAO_ATIVA and not self._timer.isActive():
                self._timer.start()
        elif watched is self.parentWidget() and event.type() == QEvent.Hide:
            # Evita gastar processamento enquanto a janela está oculta.
            self._timer.stop()

        return super().eventFilter(watched, event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setPen(Qt.NoPen)

        largura = max(1, self.width())
        altura = max(1, self.height())
        menor_lado = min(largura, altura)

        # Base clara para preservar o contraste dos textos e controles.
        painter.fillRect(self.rect(), QColor(CORES["fundo"]))

        fase = self._fase if ANIMACAO_ATIVA else 0.0
        escala = menor_lado * TAMANHO_BLUR

        # (cor, frequência X, fase X, frequência Y, fase Y, raio, intensidade)
        trajetorias = [
            (CORES["coral"],      0.59, 2.1, 0.67, 4.2, escala * 0.72, 0.92),
            (CORES["verde_agua"], 0.47, 4.0, 0.61, 2.6, escala * 0.82, 1.00),
        ]

        # Cada mancha percorre a tela tanto na horizontal quanto na vertical.
        manchas = [
            (
                cor,
                0.5 + AMPLITUDE_X * math.sin(fase * fx + px),
                0.5 + AMPLITUDE_Y * math.cos(fase * fy + py),
                raio,
                intensidade,
            )
            for cor, fx, px, fy, py, raio, intensidade in trajetorias
        ]

        for cor_hex, pos_x, pos_y, raio, intensidade in manchas:
            centro = QPointF(largura * pos_x, altura * pos_y)
            gradiente = QRadialGradient(centro, raio)

            cor_centro = QColor(cor_hex)
            cor_centro.setAlpha(round(INTENSIDADE_BLUR * intensidade))

            cor_meio = QColor(cor_hex)
            cor_meio.setAlpha(round(INTENSIDADE_BLUR * intensidade * 0.58))

            cor_transicao = QColor(cor_hex)
            cor_transicao.setAlpha(round(INTENSIDADE_BLUR * intensidade * 0.10))

            cor_borda = QColor(cor_hex)
            cor_borda.setAlpha(0)

            gradiente.setColorAt(0.00, cor_centro)
            gradiente.setColorAt(0.42, cor_meio)
            gradiente.setColorAt(0.78, cor_transicao)
            gradiente.setColorAt(1.00, cor_borda)

            painter.setBrush(gradiente)
            painter.drawEllipse(centro, raio, raio)

        painter.end()


def aplicar_transparencia_paineis(janela: QWidget) -> None:
    """Deixa translúcidas as páginas das abas que contêm botões e tabelas."""
    for abas in janela.findChildren(QTabWidget):
        for indice in range(abas.count()):
            pagina = abas.widget(indice)
            pagina.setObjectName("paginaAbaTransparente")
            pagina.setAttribute(Qt.WA_TranslucentBackground, True)
            pagina.setAutoFillBackground(False)


def aplicar_transparencia_tabelas(janela: QWidget) -> None:
    """Deixa o fundo externo da tabela transparente sem alterar as células.

    Não instala delegate personalizado: o delegate padrão do Qt lê diretamente
    Qt.BackgroundRole e Qt.ForegroundRole do PlanilhaModel. Assim, as cores
    originais das células importadas do Excel permanecem inalteradas.
    """
    for tabela in janela.findChildren(QTableView):
        tabela.setAttribute(Qt.WA_TranslucentBackground, True)
        tabela.setAutoFillBackground(False)
        tabela.viewport().setAutoFillBackground(False)


def aplicar_sombra_botoes(janela: QWidget) -> None:
    """Aplica uma sombra discreta aos botões já criados na janela."""
    for botao in janela.findChildren(QPushButton):
        sombra = QGraphicsDropShadowEffect(botao)
        sombra.setBlurRadius(13)
        sombra.setOffset(0, 2)
        sombra.setColor(QColor(118, 177, 178, 48))
        botao.setGraphicsEffect(sombra)


def aplicar_fundo_blur(janela: QWidget) -> FundoBlur:
    """Adiciona o fundo colorido desfocado à janela principal.

    Chame depois de criar a JanelaPrincipal e antes de janela.show().
    """
    if not isinstance(janela, QWidget):
        raise TypeError("aplicar_fundo_blur espera um QWidget como janela.")

    janela.setObjectName("janelaPrincipal")
    janela.setAutoFillBackground(False)

    # Guarda a referência para a camada continuar viva durante toda a aplicação.
    fundo = FundoBlur(janela)
    janela._fundo_blur = fundo
    fundo.lower()

    # Os botões e tabelas já foram criados no __init__ da janela principal.
    aplicar_sombra_botoes(janela)
    aplicar_transparencia_paineis(janela)
    aplicar_transparencia_tabelas(janela)

    return fundo


# ============================================================
# POP-UPS BRANCOS TRANSLÚCIDOS
# ============================================================

ESTILO_POPUP_BRANCO = f"""
QDialog, QMessageBox, QInputDialog {{
    background-color: rgba(255, 255, 255, {OPACIDADE_POPUP});
    border: 1px solid rgba(255, 255, 255, 230);
    border-radius: 14px;
}}

QLabel {{
    background-color: transparent;
    color: {CORES['texto']};
    font-size: 13px;
}}

QLabel#qt_msgbox_label {{
    font-size: 14px;
    font-weight: 600;
}}

QLineEdit, QComboBox {{
    background-color: rgba(255, 255, 255, 235);
    color: {CORES['texto']};
    border: 1px solid {CORES['borda']};
    border-radius: 6px;
    padding: 6px 9px;
}}

QLineEdit:focus, QComboBox:focus {{
    border: 1px solid {CORES['verde_agua']};
}}

QComboBox QAbstractItemView {{
    background-color: {CORES['superficie']};
    color: {CORES['texto']};
    selection-background-color: {CORES['selecao']};
    selection-color: {CORES['selecao_texto']};
}}

QPushButton {{
    background-color: rgba(255, 255, 255, 235);
    color: {CORES['primaria']};
    border: 1px solid rgba(118, 177, 178, 150);
    border-radius: 7px;
    padding: 7px 18px;
    min-width: 72px;
    font-weight: 600;
}}

QPushButton:hover {{
    background-color: rgba(118, 177, 178, 95);
    border: 1px solid rgba(118, 177, 178, 200);
}}

QPushButton:pressed {{
    background-color: rgba(118, 177, 178, 165);
}}

QPushButton:default {{
    border: 1px solid {CORES['verde_agua']};
}}
"""

ESTILO_PROGRESSO_BRANCO = f"""
QDialog {{
    background: transparent;
}}

QFrame#cartaoProgresso {{
    background-color: rgba(255, 255, 255, {OPACIDADE_POPUP_PROGRESSO});
    border: 1px solid rgba(255, 255, 255, 220);
    border-radius: 14px;
}}

QLabel {{
    background: transparent;
    color: {CORES['texto']};
}}

QLabel#tituloProgresso {{
    font-size: 15px;
    font-weight: 600;
    color: {CORES['primaria']};
}}

QLabel#textoProgresso {{
    font-size: 12px;
    color: {CORES['texto_secundario']};
}}

QProgressBar {{
    background-color: rgba(0, 0, 0, 35);
    border: none;
    border-radius: 1px;
    min-height: 3px;
    max-height: 3px;
}}

QProgressBar::chunk {{
    background-color: {CORES['verde_agua']};
    border-radius: 1px;
}}
"""


def preparar_popup(dlg: QWidget) -> None:
    """Deixa um pop-up sem moldura, branco e translúcido."""
    dlg.setWindowFlags(dlg.windowFlags() | Qt.FramelessWindowHint)
    dlg.setAttribute(Qt.WA_TranslucentBackground, True)
    dlg.setStyleSheet(ESTILO_POPUP_BRANCO)


class QMessageBox(_QMessageBox):
    """Mesmas chamadas estáticas do QMessageBox, com visual branco translúcido."""

    @staticmethod
    def _mostrar(parent, icone, titulo, texto):
        caixa = _QMessageBox(parent)
        caixa.setIcon(icone)
        caixa.setText(titulo)
        caixa.setInformativeText(texto)
        caixa.setStandardButtons(_QMessageBox.Ok)
        preparar_popup(caixa)
        return caixa.exec()

    @staticmethod
    def critical(parent, titulo, texto, *args, **kwargs):
        return QMessageBox._mostrar(parent, _QMessageBox.Critical, titulo, texto)

    @staticmethod
    def warning(parent, titulo, texto, *args, **kwargs):
        return QMessageBox._mostrar(parent, _QMessageBox.Warning, titulo, texto)

    @staticmethod
    def information(parent, titulo, texto, *args, **kwargs):
        return QMessageBox._mostrar(parent, _QMessageBox.Information, titulo, texto)


class QInputDialog(_QInputDialog):
    """getText e getItem com visual branco translúcido."""

    @staticmethod
    def getText(parent, titulo, rotulo, echo=QLineEdit.Normal, *args, **kwargs):
        dlg = _QInputDialog(parent)
        dlg.setInputMode(_QInputDialog.TextInput)
        dlg.setWindowTitle(titulo)
        dlg.setLabelText(rotulo)
        dlg.setTextEchoMode(echo)
        preparar_popup(dlg)
        aceito = dlg.exec() == QDialog.Accepted
        return dlg.textValue(), aceito

    @staticmethod
    def getItem(parent, titulo, rotulo, itens, atual=0, editavel=True, *args, **kwargs):
        dlg = _QInputDialog(parent)
        dlg.setInputMode(_QInputDialog.TextInput)
        dlg.setWindowTitle(titulo)
        dlg.setLabelText(rotulo)
        dlg.setComboBoxItems(list(itens))
        dlg.setComboBoxEditable(editavel)
        if 0 <= atual < len(itens):
            dlg.setTextValue(itens[atual])
        preparar_popup(dlg)
        aceito = dlg.exec() == QDialog.Accepted
        return dlg.textValue(), aceito


class DialogoProgressoLinha(QDialog):
    """Diálogo branco translúcido, sem moldura, com barra de progresso em linha fina.

    Substitui o QProgressDialog: mantém setValue, setLabelText e close.
    """

    def __init__(self, parent, titulo="", texto=""):
        super().__init__(parent)

        self.setWindowTitle(titulo)
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setWindowModality(Qt.WindowModal)
        self.setMinimumWidth(440)

        self.cartao = QFrame(self)
        self.cartao.setObjectName("cartaoProgresso")

        self.label_titulo = QLabel(titulo)
        self.label_titulo.setObjectName("tituloProgresso")

        self.label_texto = QLabel(texto)
        self.label_texto.setObjectName("textoProgresso")
        self.label_texto.setWordWrap(True)

        self.barra = QProgressBar()
        self.barra.setRange(0, 100)
        self.barra.setValue(0)
        self.barra.setTextVisible(False)
        self.barra.setFixedHeight(3)

        interno = QVBoxLayout(self.cartao)
        interno.setContentsMargins(24, 20, 24, 22)
        interno.setSpacing(12)
        interno.addWidget(self.label_titulo)
        interno.addWidget(self.label_texto)
        interno.addSpacing(4)
        interno.addWidget(self.barra)

        externo = QVBoxLayout(self)
        externo.setContentsMargins(0, 0, 0, 0)
        externo.addWidget(self.cartao)

        self.setStyleSheet(ESTILO_PROGRESSO_BRANCO)

    # API compatível com QProgressDialog
    def setValue(self, valor):
        self.barra.setValue(int(valor))

    def setLabelText(self, texto):
        self.label_texto.setText(texto)

    def showEvent(self, event):
        super().showEvent(event)
        pai = self.parentWidget()
        if pai is not None:
            centro = pai.frameGeometry().center()
            geo = self.frameGeometry()
            geo.moveCenter(centro)
            self.move(geo.topLeft())

    def keyPressEvent(self, event):
        # Impede fechar com ESC durante o processamento.
        if event.key() == Qt.Key_Escape:
            event.ignore()
            return
        super().keyPressEvent(event)


# ============================================================
# STYLE SHEET (QSS)
# ============================================================

ESTILO_APP = f"""
/* O fundo da janela é desenhado pela classe FundoBlur. */
QWidget {{
    color: {CORES['texto']};
    font-family: "{FONTE_FAMILIA}";
    font-size: {FONTE_TAMANHO}pt;
}}

QWidget#janelaPrincipal {{
    background-color: {CORES['fundo']};
}}

QWidget#fundoBlur {{
    background-color: transparent;
}}

/* Abas */
QTabWidget::pane {{
    background-color: rgba(243, 246, 249, 135);
    border: 1px solid rgba(217, 225, 232, 115);
    border-radius: 7px;
    top: -1px;
}}

/* Página interna das abas: painel atrás dos botões e da tabela. */
QWidget#paginaAbaTransparente {{
    background-color: rgba(243, 246, 249, 105);
    border: none;
}}

QTabBar::tab {{
    /* Abas não selecionadas sem preenchimento: o blur fica visível. */
    background-color: transparent;
    color: {CORES['texto_secundario']};
    border: 1px solid transparent;
    border-bottom: none;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    padding: 10px 22px;
    margin-right: 3px;
    min-width: 90px;
}}

QTabBar::tab:hover {{
    /* Hover sem fundo para manter a aba não selecionada transparente. */
    background-color: transparent;
    color: {CORES['primaria']};
    border-color: rgba(118, 177, 178, 75);
}}

QTabBar::tab:selected {{
    /* Aba ativa com branco translúcido. */
    background-color: rgba(255, 255, 255, {OPACIDADE_ABA_SELECIONADA});
    color: {CORES['primaria']};
    border: 1px solid rgba(255, 255, 255, 185);
    border-bottom: none;
    font-weight: 600;
}}

QTabBar::tab:selected:hover {{
    background-color: rgba(255, 255, 255, {OPACIDADE_ABA_SELECIONADA});
}}

/* Botões transparentes com sombra discreta aplicada via QGraphicsDropShadowEffect. */
QPushButton {{
    background-color: rgba(255, 255, 255, 38);
    color: {CORES['primaria']};
    border: 1px solid rgba(118, 177, 178, 110);
    border-radius: 7px;
    padding: 9px 16px;
    min-height: 18px;
    font-weight: 600;
}}

QPushButton:hover {{
    background-color: rgba(118, 177, 178, 95);
    color: {CORES['primaria']};
    border: 1px solid rgba(118, 177, 178, 190);
}}

QPushButton:pressed {{
    background-color: rgba(118, 177, 178, 165);
    color: {CORES['primaria']};
    border: 1px solid rgba(118, 177, 178, 225);
    padding-top: 10px;
    padding-bottom: 8px;
}}

QPushButton:checked {{
    background-color: rgba(118, 177, 178, 125);
    color: {CORES['primaria']};
    border: 1px solid rgba(118, 177, 178, 220);
}}

QPushButton:disabled {{
    background-color: rgba(170, 183, 196, 55);
    color: rgba(38, 52, 69, 115);
    border: 1px solid rgba(170, 183, 196, 90);
}}

/* Tabelas: o fundo é transparente, mas as cores das células vêm do modelo. */
QTableView {{
    background-color: transparent;
    alternate-background-color: transparent;
    color: {CORES['texto']};
    gridline-color: {CORES['borda']};
    border: 1px solid {CORES['borda']};
    border-radius: 6px;
    selection-background-color: {CORES['selecao']};
    selection-color: {CORES['selecao_texto']};
    outline: 0;
}}

QTableView::item {{
    padding: 4px 6px;
    border: none;
}}

QTableView::item:selected {{
    background-color: {CORES['selecao']};
    color: {CORES['selecao_texto']};
}}

/* Cabeçalho vertical: faixa dos números das linhas em branco. */
QHeaderView::section:vertical {{
    background-color: #FFFFFF;
    color: #0E3862;
    border: none;
    border-right: 1px solid #D9E1E8;
    border-bottom: 1px solid #D9E1E8;
    padding: 4px 6px;
    font-weight: 500;
}}

/* Canto superior esquerdo da tabela, acima dos números das linhas. */
QTableCornerButton::section {{
    background-color: #FFFFFF;
    border: none;
    border-right: 1px solid #D9E1E8;
    border-bottom: 1px solid #D9E1E8;
}}

/*
Não impor cores personalizadas aos cabeçalhos ou ao botão de canto da tabela.
Eles usam o estilo padrão do Qt/Fusion, sem o azul imposto pelo tema.
*/

/* Campos de texto e seleção */
QLineEdit, QTextEdit, QPlainTextEdit, QComboBox, QSpinBox, QDoubleSpinBox {{
    background-color: rgba(255, 255, 255, 240);
    color: {CORES['texto']};
    border: 1px solid {CORES['borda']};
    border-radius: 5px;
    padding: 7px 9px;
    selection-background-color: {CORES['selecao']};
    selection-color: {CORES['selecao_texto']};
}}

QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus,
QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus {{
    border: 1px solid {CORES['verde_agua']};
}}

QComboBox::drop-down {{
    border: none;
    width: 25px;
}}

QComboBox QAbstractItemView {{
    background-color: {CORES['superficie']};
    color: {CORES['texto']};
    border: 1px solid {CORES['borda']};
    selection-background-color: {CORES['selecao']};
    selection-color: {CORES['selecao_texto']};
}}

/* Rótulos */
QLabel {{
    background-color: transparent;
    color: {CORES['texto']};
}}

/* Barras de rolagem */
QScrollBar:vertical {{
    background-color: transparent;
    width: 12px;
    margin: 2px;
    border: none;
}}

QScrollBar::handle:vertical {{
    background-color: #B8C5D1;
    min-height: 28px;
    border-radius: 5px;
}}

QScrollBar::handle:vertical:hover {{
    background-color: {CORES['verde_agua']};
}}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
    border: none;
}}

QScrollBar:horizontal {{
    background-color: transparent;
    height: 12px;
    margin: 2px;
    border: none;
}}

QScrollBar::handle:horizontal {{
    background-color: #B8C5D1;
    min-width: 28px;
    border-radius: 5px;
}}

QScrollBar::handle:horizontal:hover {{
    background-color: {CORES['verde_agua']};
}}

QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0;
    border: none;
}}

/*
Pop-ups de mensagem, entrada e progresso são estilizados em branco translúcido
por preparar_popup() e DialogoProgressoLinha. O seletor de arquivos é nativo
do sistema e mantém o fundo sólido.
*/
QFileDialog {{
    background-color: {CORES['fundo']};
}}

QProgressBar {{
    background-color: {CORES['superficie_secundaria']};
    color: {CORES['texto']};
    border: 1px solid {CORES['borda']};
    border-radius: 5px;
    text-align: center;
    min-height: 16px;
}}

QProgressBar::chunk {{
    background-color: {CORES['verde_agua']};
    border-radius: 4px;
}}

/* Menus */
QMenu {{
    background-color: {CORES['superficie']};
    color: {CORES['texto']};
    border: 1px solid {CORES['borda']};
    padding: 5px;
}}

QMenu::item {{
    padding: 7px 24px;
    border-radius: 3px;
}}

QMenu::item:selected {{
    background-color: {CORES['selecao']};
    color: {CORES['selecao_texto']};
}}
"""


# ============================================================
# APLICAÇÃO DO ESTILO GLOBAL
# ============================================================

def aplicar_estilo(app: QApplication) -> None:
    """Aplica fontes, paleta e QSS à instância QApplication."""
    if not isinstance(app, QApplication):
        raise TypeError("aplicar_estilo espera uma instância de QApplication.")

    # Fusion ajuda a manter uma aparência mais consistente entre sistemas.
    app.setStyle("Fusion")
    app.setFont(QFont(FONTE_FAMILIA, FONTE_TAMANHO))

    paleta = app.palette()
    paleta.setColor(QPalette.Window, QColor(CORES["fundo"]))
    paleta.setColor(QPalette.WindowText, QColor(CORES["texto"]))
    paleta.setColor(QPalette.Base, QColor(CORES["superficie"]))
    paleta.setColor(QPalette.AlternateBase, QColor(CORES["alternada"]))
    paleta.setColor(QPalette.Text, QColor(CORES["texto"]))
    paleta.setColor(QPalette.Button, QColor(CORES["primaria"]))
    paleta.setColor(QPalette.ButtonText, QColor(CORES["branco"]))
    paleta.setColor(QPalette.Highlight, QColor(CORES["selecao"]))
    paleta.setColor(QPalette.HighlightedText, QColor(CORES["selecao_texto"]))
    app.setPalette(paleta)
    app.setStyleSheet(ESTILO_APP)
