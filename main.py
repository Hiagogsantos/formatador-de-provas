# -*- coding: utf-8 -*-
"""
main.py
-------
Ponto de entrada do Formatador de Provas.

Requisitos (instalar antes de rodar):
    pip install PyQt6 python-docx pillow lxml
    pip install pymupdf          # opcional, só para a pré-visualização

Além disso, para a pré-visualização funcionar, é preciso ter o LibreOffice
instalado no computador (o mesmo já usado para gerar PDFs a partir de
.docx). Sem ele, o app formata normalmente — só a miniatura de preview
não aparece.

Para rodar:
    python main.py
"""

import sys

from PyQt6.QtWidgets import QApplication

from ui.styles import STYLESHEET
from ui.janela_principal import JanelaPrincipal


def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)

    janela = JanelaPrincipal()
    janela.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
