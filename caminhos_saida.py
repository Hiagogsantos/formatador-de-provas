# -*- coding: utf-8 -*-
"""
caminhos_saida.py
------------------
Resolve onde salvar a prova formatada:

    provas/<ANO>/<BIMESTRE EXTENSO>/<MATÉRIA>/<nome do arquivo>.docx

Ex.: provas/2026/III BIMESTRE/PORTUGUÊS/PROVA_PORTUGUES_2ANO_A.docx

Cria as pastas automaticamente se não existirem. Se já existir um arquivo
com o mesmo nome, adiciona um sufixo numérico (_2, _3, ...) para nunca
sobrescrever uma prova sem querer.
"""

import re
import unicodedata
from pathlib import Path


def _sanitizar_nome_pasta(texto: str) -> str:
    """Remove caracteres que o Windows não aceita em nome de pasta/arquivo
    (: / \\ * ? " < > |), mantendo acentos (o Windows aceita acentos)."""
    texto = texto.strip()
    return re.sub(r'[\\/:*?"<>|]', "", texto)


def _sanitizar_nome_arquivo(texto: str) -> str:
    base = _sanitizar_nome_pasta(texto)
    base = re.sub(r"\s+", "_", base)
    return base


def resolver_pasta_destino(pasta_provas: Path, ano: str, bimestre_extenso: str, materia: str) -> Path:
    """Monta e cria (se preciso) a pasta provas/<ano>/<bimestre>/<materia>."""
    pasta = (
        pasta_provas
        / _sanitizar_nome_pasta(str(ano))
        / _sanitizar_nome_pasta(bimestre_extenso)
        / _sanitizar_nome_pasta(materia)
    )
    pasta.mkdir(parents=True, exist_ok=True)
    return pasta


def gerar_caminho_saida(pasta_provas: Path, ano: str, bimestre_extenso: str,
                         materia: str, nome_base_arquivo: str) -> Path:
    """
    Retorna o Path completo do .docx de saída, já dentro da pasta correta
    e sem colidir com um arquivo existente (adiciona _2, _3... se precisar).

    `nome_base_arquivo` normalmente vem do nome do arquivo original (sem
    extensão) ou de um nome digitado pelo usuário.
    """
    pasta_destino = resolver_pasta_destino(pasta_provas, ano, bimestre_extenso, materia)
    nome_limpo = _sanitizar_nome_arquivo(nome_base_arquivo) or "PROVA_FORMATADA"

    candidato = pasta_destino / f"{nome_limpo}.docx"
    contador = 2
    while candidato.exists():
        candidato = pasta_destino / f"{nome_limpo}_{contador}.docx"
        contador += 1
    return candidato
