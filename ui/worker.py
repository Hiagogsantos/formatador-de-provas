# -*- coding: utf-8 -*-
"""
ui/worker.py
-------------
Roda a formatação da prova (e a geração do preview) numa QThread separada,
para a janela não travar/"congelar" enquanto o python-docx e o LibreOffice
processam o arquivo (isso pode levar alguns segundos).
"""

import subprocess
import sys
from pathlib import Path

from PyQt6.QtCore import QThread, pyqtSignal

import formatar_prova


class WorkerFormatacao(QThread):
    """
    Parâmetros esperados em `dados` (dict):
        entrada, saida, disciplina, bimestre, modelo_cabecalho,
        valores_cabecalho (dict)

    Sinais:
        sucesso(str caminho_docx, str caminho_pdf_preview_ou_vazio)
        erro(str mensagem)
    """

    sucesso = pyqtSignal(str, str)
    erro = pyqtSignal(str)

    def __init__(self, dados: dict, parent=None):
        super().__init__(parent)
        self.dados = dados

    def run(self):
        try:
            formatar_prova.formatar_prova(
                self.dados["entrada"],
                self.dados["saida"],
                disciplina=self.dados.get("disciplina"),
                bimestre=self.dados.get("bimestre", "3º"),
                caminho_docx_modelo_cabecalho=self.dados.get("modelo_cabecalho"),
                valores_cabecalho=self.dados.get("valores_cabecalho"),
            )
        except Exception as exc:  # noqa: BLE001 - queremos capturar tudo e mostrar na UI
            self.erro.emit(f"Falha ao formatar a prova:\n{exc}")
            return

        caminho_pdf = self._gerar_preview_pdf(self.dados["saida"])
        self.sucesso.emit(self.dados["saida"], caminho_pdf or "")

    def _gerar_preview_pdf(self, caminho_docx: str) -> str | None:
        """
        Converte o .docx recém-gerado para PDF usando o LibreOffice
        instalado no sistema (soffice), para servir de preview da primeira
        página dentro do app. Se o LibreOffice não estiver instalado, o
        preview simplesmente não aparece — a formatação em si já foi
        concluída com sucesso de qualquer forma.
        """
        docx_path = Path(caminho_docx)
        pasta_saida = docx_path.parent

        candidatos_soffice = ["soffice", "soffice.exe"]
        executavel = None
        for nome in candidatos_soffice:
            try:
                subprocess.run(
                    [nome, "--version"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    timeout=10,
                )
                executavel = nome
                break
            except (FileNotFoundError, subprocess.TimeoutExpired):
                continue

        if executavel is None:
            return None

        try:
            subprocess.run(
                [
                    executavel,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(pasta_saida),
                    str(docx_path),
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=60,
                check=True,
            )
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
            return None

        caminho_pdf = pasta_saida / (docx_path.stem + ".pdf")
        return str(caminho_pdf) if caminho_pdf.exists() else None
