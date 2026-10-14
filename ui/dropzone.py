# -*- coding: utf-8 -*-
"""
ui/dropzone.py
---------------
Área de "arrastar e soltar" o .docx da prova crua. Também responde a
clique simples, abrindo o explorador de arquivos do Windows para escolher
manualmente. Emite o sinal `arquivo_selecionado(str)` com o caminho
completo assim que um .docx válido é solto ou escolhido.
"""

from pathlib import Path

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLabel, QFileDialog

from ui.styles import repolir

EXTENSOES_ACEITAS = (".docx",)


class DropZone(QWidget):
    """Clique ou arraste um .docx para cá. Estados visuais (via
    setProperty("estado", ...)):
        "vazio"       -> nada selecionado ainda
        "arrastando"  -> um arquivo está sendo arrastado por cima agora
        "preenchido"  -> já tem um .docx válido selecionado
    """

    arquivo_selecionado = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setProperty("cls", "dropzone")
        self.setAcceptDrops(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(150)

        self._caminho_atual = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(6)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.rotulo_icone = QLabel("📄")
        self.rotulo_icone.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.rotulo_icone.setStyleSheet("font-size: 30px; background: transparent;")
        layout.addWidget(self.rotulo_icone)

        self.rotulo_titulo = QLabel("Arraste a prova (.docx) aqui")
        self.rotulo_titulo.setProperty("cls", "dropzoneTitulo")
        self.rotulo_titulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.rotulo_titulo)

        self.rotulo_subtitulo = QLabel("ou clique para escolher no computador")
        self.rotulo_subtitulo.setProperty("cls", "dropzoneSubtitulo")
        self.rotulo_subtitulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.rotulo_subtitulo)

        self.rotulo_arquivo = QLabel("")
        self.rotulo_arquivo.setProperty("cls", "dropzoneArquivo")
        self.rotulo_arquivo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.rotulo_arquivo.setWordWrap(True)
        self.rotulo_arquivo.hide()
        layout.addWidget(self.rotulo_arquivo)

        self._definir_estado("vazio")

    # ---------------------------------------------------------------
    # API pública
    # ---------------------------------------------------------------

    def caminho_atual(self):
        return self._caminho_atual

    def limpar(self):
        self._caminho_atual = None
        self.rotulo_arquivo.hide()
        self.rotulo_icone.setText("📄")
        self.rotulo_titulo.setText("Arraste a prova (.docx) aqui")
        self.rotulo_subtitulo.setText("ou clique para escolher no computador")
        self._definir_estado("vazio")

    # ---------------------------------------------------------------
    # Clique -> abre o explorador de arquivos do Windows
    # ---------------------------------------------------------------

    def mousePressEvent(self, evento):
        if evento.button() == Qt.MouseButton.LeftButton:
            self._abrir_dialogo_escolha()
        super().mousePressEvent(evento)

    def _abrir_dialogo_escolha(self):
        caminho, _ = QFileDialog.getOpenFileName(
            self,
            "Escolher prova (.docx)",
            "",
            "Documentos Word (*.docx)",
        )
        if caminho:
            self._definir_arquivo(caminho)

    # ---------------------------------------------------------------
    # Arrastar e soltar
    # ---------------------------------------------------------------

    def dragEnterEvent(self, evento):
        if self._evento_tem_docx_valido(evento):
            self._definir_estado("arrastando")
            evento.acceptProposedAction()
        else:
            evento.ignore()

    def dragLeaveEvent(self, evento):
        self._definir_estado("preenchido" if self._caminho_atual else "vazio")

    def dropEvent(self, evento):
        if self._evento_tem_docx_valido(evento):
            caminho = evento.mimeData().urls()[0].toLocalFile()
            self._definir_arquivo(caminho)
            evento.acceptProposedAction()
        else:
            self._definir_estado("vazio")
            evento.ignore()

    def _evento_tem_docx_valido(self, evento):
        if not evento.mimeData().hasUrls():
            return False
        urls = evento.mimeData().urls()
        if len(urls) != 1:
            return False
        caminho = urls[0].toLocalFile()
        return caminho.lower().endswith(EXTENSOES_ACEITAS)

    # ---------------------------------------------------------------
    # Internos
    # ---------------------------------------------------------------

    def _definir_arquivo(self, caminho):
        caminho = str(Path(caminho))
        if not caminho.lower().endswith(EXTENSOES_ACEITAS):
            return
        self._caminho_atual = caminho
        self.rotulo_icone.setText("✅")
        self.rotulo_titulo.setText("Prova selecionada")
        self.rotulo_subtitulo.setText("Clique para trocar o arquivo")
        self.rotulo_arquivo.setText(Path(caminho).name)
        self.rotulo_arquivo.show()
        self._definir_estado("preenchido")
        self.arquivo_selecionado.emit(caminho)

    def _definir_estado(self, estado):
        self.setProperty("estado", estado)
        repolir(self)
