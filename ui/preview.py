# -*- coding: utf-8 -*-
"""
ui/preview.py
--------------
Widget que mostra a primeira página do PDF (gerado a partir do .docx
formatado) como imagem, dentro da própria janela do app.

Usa PyMuPDF (fitz) para rasterizar a página, porque é uma dependência leve
e não exige Poppler instalado à parte (diferente de pdf2image). Se o
PyMuPDF não estiver disponível, o widget mostra uma mensagem discreta em
vez de travar o app — a formatação em si não depende disso.
"""

from pathlib import Path

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap, QImage
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLabel

try:
    import fitz  # PyMuPDF
    _TEM_PYMUPDF = True
except ImportError:
    _TEM_PYMUPDF = False


class PreviewPagina(QWidget):
    """Mostra a 1ª página de um PDF em miniatura. Chame `mostrar_pdf(caminho)`
    para carregar, ou `limpar()` para voltar ao estado vazio."""

    LARGURA_PREVIEW = 260

    def __init__(self, parent=None):
        super().__init__(parent)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.rotulo_imagem = QLabel()
        self.rotulo_imagem.setProperty("cls", "previewMoldura")
        self.rotulo_imagem.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.rotulo_imagem.setFixedSize(self.LARGURA_PREVIEW, int(self.LARGURA_PREVIEW * 1.414))
        self.rotulo_imagem.setScaledContents(False)
        layout.addWidget(self.rotulo_imagem)

        self.limpar()

    def limpar(self):
        self.rotulo_imagem.setPixmap(QPixmap())
        if _TEM_PYMUPDF:
            self.rotulo_imagem.setText("Pré-visualização\naparece aqui\napós formatar")
        else:
            self.rotulo_imagem.setText(
                "Pré-visualização indisponível\n"
                "(instale PyMuPDF: pip install pymupdf)\n\n"
                "O arquivo .docx já foi gerado normalmente."
            )

    def mostrar_pdf(self, caminho_pdf: str):
        if not caminho_pdf or not _TEM_PYMUPDF:
            self.limpar()
            return

        caminho = Path(caminho_pdf)
        if not caminho.exists():
            self.limpar()
            return

        try:
            documento = fitz.open(str(caminho))
            pagina = documento.load_page(0)
            escala = (self.LARGURA_PREVIEW * 2) / pagina.rect.width  # 2x para nitidez em telas HiDPI
            matriz = fitz.Matrix(escala, escala)
            pixmap_fitz = pagina.get_pixmap(matrix=matriz)
            documento.close()

            imagem_qt = QImage(
                pixmap_fitz.samples,
                pixmap_fitz.width,
                pixmap_fitz.height,
                pixmap_fitz.stride,
                QImage.Format.Format_RGB888,
            )
            pixmap_qt = QPixmap.fromImage(imagem_qt).scaled(
                self.rotulo_imagem.width(),
                self.rotulo_imagem.height(),
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
            self.rotulo_imagem.setText("")
            self.rotulo_imagem.setPixmap(pixmap_qt)
        except Exception:
            self.limpar()
