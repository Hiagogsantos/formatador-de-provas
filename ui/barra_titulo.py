# -*- coding: utf-8 -*-
"""
ui/barra_titulo.py
-------------------
Barra de título customizada (minimizar + fechar), copiada tal como
documentada em IDENTIDADE_VISUAL_PYQT6.md. Nunca existiu um terceiro
botão de maximizar nesse padrão — a janela principal deste app é de
tamanho fixo.
"""

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QWidget, QPushButton, QLabel, QHBoxLayout


class BarraTitulo(QWidget):
    """Barra de título customizada: minimizar + fechar. Para diálogos
    (janelas modais), use mostrar_minimizar=False — fica só o fechar."""

    ALTURA = 34

    def __init__(self, janela, titulo, mostrar_minimizar=True, parent=None):
        super().__init__(parent)
        self.janela = janela
        self._pos_arraste = None

        # sem isso, um QWidget usado como filho (não como janela de
        # topo) ignora o background-color do QSS -- é a pegadinha nº1
        # de quem tenta colorir uma barra de título customizada
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)

        self.setFixedHeight(self.ALTURA)
        self.setProperty("cls", "barraTitulo")

        lay = QHBoxLayout(self)
        lay.setContentsMargins(14, 0, 10, 0)
        lay.setSpacing(6)

        self.rotulo_titulo = QLabel(titulo)
        self.rotulo_titulo.setProperty("cls", "barraTituloTexto")
        lay.addWidget(self.rotulo_titulo)
        lay.addStretch(1)

        if mostrar_minimizar:
            btn_min = self._criar_botao("—", self.janela.showMinimized)
            lay.addWidget(btn_min, alignment=Qt.AlignmentFlag.AlignVCenter)

        btn_fechar = self._criar_botao("✕", self.janela.close)
        lay.addWidget(btn_fechar, alignment=Qt.AlignmentFlag.AlignVCenter)

    def _criar_botao(self, texto, callback):
        btn = QPushButton(texto)
        btn.setProperty("cls", "barraTituloBotao")
        btn.setFixedSize(26, 26)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.clicked.connect(callback)
        return btn

    # ---- sem a barra nativa, arrastar a janela também precisa ser
    # reimplementado na mão: guarda o offset do clique, e move a
    # janela junto com o mouse enquanto o botão estiver pressionado
    def mousePressEvent(self, evento):
        if evento.button() == Qt.MouseButton.LeftButton:
            self._pos_arraste = (
                evento.globalPosition().toPoint() - self.janela.frameGeometry().topLeft()
            )
            evento.accept()

    def mouseMoveEvent(self, evento):
        if self._pos_arraste is not None and evento.buttons() & Qt.MouseButton.LeftButton:
            self.janela.move(evento.globalPosition().toPoint() - self._pos_arraste)
            evento.accept()

    def mouseReleaseEvent(self, evento):
        self._pos_arraste = None
