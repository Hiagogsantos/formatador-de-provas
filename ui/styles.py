# -*- coding: utf-8 -*-
"""
ui/styles.py
------------
Cores + QSS da identidade visual da escola (ver IDENTIDADE_VISUAL_PYQT6.md).
Aplicado uma vez via app.setStyleSheet(STYLESHEET) no main.py.

Segue exatamente o padrão dos outros projetos:
  - cores em Python (fácil reaproveitar em código, ex.: QColor(LARANJA))
  - seleção de estilo por propriedade "cls" (setProperty("cls", "..."))
  - layout NUNCA fica aqui — isso é sempre QVBoxLayout/QHBoxLayout no
    código Python das telas
"""

# --------------------------------------------------------------------------
# Cores
# --------------------------------------------------------------------------

LARANJA = "#db6018"
LARANJA_HOVER = "#ea6c1f"
LARANJA_LINHA = "rgba(219, 96, 24, 140)"
LARANJA_FUNDO = "rgba(219, 96, 24, 35)"
PRETO_FUNDO = "#0c0c0c"
PRETO_PAINEL = "#131313"
PRETO_CAMPO = "#181818"
LINHA = "#262421"
LINHA_SUAVE = "#1c1a18"
TEXTO = "#efece6"
TEXTO_DIM = "#9a958c"
TEXTO_FRACO = "#5f5b54"

VERDE_SUCESSO = "#4caf7d"
VERMELHO_ERRO = "#e0554f"


def repolir(widget):
    """Força o Qt a reaplicar o QSS depois de mudar uma propriedade
    dinâmica (ex.: setProperty("estado", "arrastando")) em runtime."""
    widget.style().unpolish(widget)
    widget.style().polish(widget)


# --------------------------------------------------------------------------
# QSS
# --------------------------------------------------------------------------

STYLESHEET = f"""
QWidget {{
    color: {TEXTO};
    font-family: "Segoe UI";
    font-size: 12px;
    background: transparent;
}}
QMainWindow, QDialog {{
    background-color: {PRETO_FUNDO};
}}

/* ---------- Barra de título customizada ---------- */
QWidget[cls="barraTitulo"] {{
    background-color: {PRETO_CAMPO};
}}
QLabel[cls="barraTituloTexto"] {{
    color: {LARANJA};
    font-size: 11px;
    font-weight: 600;
}}
QPushButton[cls="barraTituloBotao"] {{
    background-color: {PRETO_FUNDO};
    border: 1px solid {TEXTO_DIM};
    border-radius: 6px;
    padding: 0px;
    color: {TEXTO_DIM};
    font-size: 12px;
}}
QPushButton[cls="barraTituloBotao"]:hover {{
    background-color: {LARANJA};
    border: 1px solid #140b04;
    color: #140b04;
}}

/* ---------- Botões padrão ---------- */
QPushButton {{
    background-color: {PRETO_CAMPO};
    border: 1px solid {LINHA};
    border-radius: 7px;
    padding: 8px 16px;
    color: {TEXTO};
    font-weight: 600;
}}
QPushButton:hover {{ border-color: {TEXTO_DIM}; }}
QPushButton:disabled {{ color: {TEXTO_FRACO}; border-color: {LINHA_SUAVE}; }}
QPushButton[variant="primary"] {{
    background-color: {LARANJA};
    border-color: {LARANJA};
    color: #140b04;
}}
QPushButton[variant="primary"]:hover {{ background-color: {LARANJA_HOVER}; }}
QPushButton[variant="primary"]:disabled {{
    background-color: {LINHA_SUAVE};
    border-color: {LINHA_SUAVE};
    color: {TEXTO_FRACO};
}}

/* ---------- Campos de texto / combos ---------- */
QLineEdit, QComboBox, QSpinBox {{
    background-color: {PRETO_CAMPO};
    border: 1px solid {LINHA};
    border-radius: 6px;
    padding: 6px 8px;
    color: {TEXTO};
    selection-background-color: {LARANJA};
}}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus {{
    border: 1px solid {LARANJA};
}}
QComboBox::drop-down {{
    border: none;
    width: 22px;
}}
QComboBox QAbstractItemView {{
    background-color: {PRETO_CAMPO};
    border: 1px solid {LINHA};
    selection-background-color: {LARANJA_FUNDO};
    selection-color: {TEXTO};
    outline: none;
    padding: 4px;
}}

/* ---------- Rótulos de seção (título de cada bloco do form) ---------- */
QLabel[cls="rotuloSecao"] {{
    color: {LARANJA};
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.5px;
}}
QLabel[cls="rotuloCampo"] {{
    color: {TEXTO_DIM};
    font-size: 11px;
    font-weight: 600;
}}

/* ---------- Painel / caixa genérica ---------- */
QWidget[cls="painel"] {{
    background-color: {PRETO_PAINEL};
    border: 1px solid {LINHA};
    border-radius: 10px;
}}

/* ---------- Área de arrastar-e-soltar (dropzone) ---------- */
QWidget[cls="dropzone"] {{
    background-color: {PRETO_PAINEL};
    border: 2px dashed {LINHA};
    border-radius: 10px;
}}
QWidget[cls="dropzone"][estado="arrastando"] {{
    background-color: {LARANJA_FUNDO};
    border: 2px dashed {LARANJA};
}}
QWidget[cls="dropzone"][estado="preenchido"] {{
    background-color: {PRETO_PAINEL};
    border: 2px solid {LARANJA};
}}
QLabel[cls="dropzoneTitulo"] {{
    color: {TEXTO};
    font-size: 14px;
    font-weight: 700;
}}
QLabel[cls="dropzoneSubtitulo"] {{
    color: {TEXTO_DIM};
    font-size: 11px;
}}
QLabel[cls="dropzoneArquivo"] {{
    color: {LARANJA};
    font-size: 12px;
    font-weight: 600;
}}

/* ---------- Barra de progresso / status ---------- */
QProgressBar {{
    background-color: {PRETO_CAMPO};
    border: 1px solid {LINHA};
    border-radius: 6px;
    text-align: center;
    color: {TEXTO};
    height: 18px;
}}
QProgressBar::chunk {{
    background-color: {LARANJA};
    border-radius: 5px;
}}

QLabel[cls="status"] {{
    font-size: 11px;
    color: {TEXTO_DIM};
}}
QLabel[cls="statusSucesso"] {{
    font-size: 12px;
    font-weight: 700;
    color: {VERDE_SUCESSO};
}}
QLabel[cls="statusErro"] {{
    font-size: 12px;
    font-weight: 700;
    color: {VERMELHO_ERRO};
}}

/* ---------- Preview de página (imagem dentro de moldura) ---------- */
QLabel[cls="previewMoldura"] {{
    background-color: {PRETO_CAMPO};
    border: 1px solid {LINHA};
    border-radius: 8px;
}}

/* ---------- Scrollbars discretas ---------- */
QScrollBar:vertical {{
    background: transparent;
    width: 10px;
    margin: 0px;
}}
QScrollBar::handle:vertical {{
    background: {LINHA};
    border-radius: 5px;
    min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{ background: {TEXTO_FRACO}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
"""
