# -*- coding: utf-8 -*-
"""
config.py
---------
Constantes centrais do app: lista de matérias, bimestres, caminhos de
pastas, e nome do arquivo de frases motivacionais.

Mantido separado do resto para você poder ajustar rapidamente (ex.: trocar
a lista de matérias, ou o caminho da pasta de saída) sem mexer na lógica.
"""

from pathlib import Path

# --------------------------------------------------------------------------
# Pastas
# --------------------------------------------------------------------------

# Pasta base do projeto (onde este arquivo está)
PASTA_BASE = Path(__file__).resolve().parent

# Todas as provas formatadas são salvas dentro de PASTA_PROVAS, organizadas
# como: provas/<ANO>/<BIMESTRE>/<MATERIA>/arquivo.docx
# Ex.: provas/2026/III BIMESTRE/PORTUGUÊS/PROVA_PORTUGUES_2ANO.docx
PASTA_PROVAS = PASTA_BASE / "provas"

CAMINHO_FRASES_JSON = PASTA_BASE / "frases_motivacionais.json"

# .docx de referência de onde puxar o cabeçalho decorado (logo + campos).
# Pode apontar pra qualquer prova já diagramada no padrão da escola. Deixe
# None se ainda não tiver um modelo salvo — a interface também permite
# escolher um modelo diferente a cada execução.
CAMINHO_MODELO_CABECALHO_PADRAO = PASTA_BASE / "modelo_cabecalho.docx"


# --------------------------------------------------------------------------
# Matérias (8 no total, conforme pedido)
# --------------------------------------------------------------------------

MATERIAS = [
    "PORTUGUÊS",
    "MATEMÁTICA",
    "CIÊNCIAS",
    "HISTÓRIA",
    "GEOGRAFIA",
    "INGLÊS",
    "ARTES",
    "ENSINO RELIGIOSO",
    "EDUCAÇÃO FÍSICA",
]


# --------------------------------------------------------------------------
# Bimestres — dois formatos:
#   - "curto": usado no título da caixa laranja ("...– 3º BIMESTRE")
#   - "extenso": usado no campo "Bimestre" do cabeçalho decorado
#     ("III BIMESTRE")
# --------------------------------------------------------------------------

BIMESTRES = [
    {"curto": "1º", "extenso": "I BIMESTRE"},
    {"curto": "2º", "extenso": "II BIMESTRE"},
    {"curto": "3º", "extenso": "III BIMESTRE"},
    {"curto": "4º", "extenso": "IV BIMESTRE"},
]

# Nome de pasta amigável para cada bimestre (usado em PASTA_PROVAS/<ano>/<isso>/<materia>)
def nome_pasta_bimestre(bimestre_extenso: str) -> str:
    return bimestre_extenso  # ex.: "III BIMESTRE" já serve como nome de pasta


# --------------------------------------------------------------------------
# Identidade visual (extraída de IDENTIDADE_VISUAL_PYQT6.md)
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
