# -*- coding: utf-8 -*-
"""
frases.py
---------
Carrega frases_motivacionais.json e expõe a lista de frases de uma
matéria específica. Recarrega o arquivo do disco a cada chamada, então
editar o .json e reabrir o app (ou trocar de matéria e voltar) já reflete
as mudanças sem precisar mexer em código.
"""

import json
from pathlib import Path

FRASE_PADRAO = "SUA DEDICAÇÃO É O QUE TE FAZ VOAR LONGE!"


def carregar_frases(caminho_json: Path) -> dict:
    """Retorna o dict {materia: [frases...]}. Se o arquivo não existir ou
    estiver malformado, retorna {} silenciosamente (a interface cai no
    fallback de frase padrão nesse caso)."""
    try:
        with open(caminho_json, encoding="utf-8") as f:
            dados = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}
    return {k: v for k, v in dados.items() if not k.startswith("_")}


def frases_da_materia(caminho_json: Path, materia: str) -> list:
    """Lista de frases cadastradas para `materia`. Se a matéria não tiver
    frases cadastradas, devolve uma lista só com a frase padrão (garante
    que o dropdown nunca fique vazio)."""
    todas = carregar_frases(caminho_json)
    frases = todas.get(materia, [])
    return frases if frases else [FRASE_PADRAO]
