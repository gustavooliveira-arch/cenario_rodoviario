#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Personalização visual da aplicação PySide6.

Integração no main.py:
    from estilo_app import aplicar_estilo, aplicar_fundo_blur

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
"""

import math

from PySide6.QtCore import QEvent, QPointF, Qt, QTimer
from PySide6.QtGui import QFont, QPalette, QColor, QPainter, QRadialGradient
from PySide6.QtWidgets import (
    QApplication,
    QGraphicsDropShadowEffect,
    QPushButton,
    QTableView,
    QTabWidget,
    QWidget,
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

/* Caixas de diálogo e progresso */
QMessageBox, QProgressDialog, QInputDialog, QFileDialog {{
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
