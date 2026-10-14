# -*- coding: utf-8 -*-
"""
ui/janela_principal.py
------------------------
Janela principal do app Formatador de Provas.

Layout geral (duas colunas dentro de um painel único):
    esquerda -> formulário (matéria, bimestre, professor, turma, data,
                frase motivacional) + dropzone do arquivo + botão formatar
    direita  -> preview da primeira página + status/mensagens

Sem barra de menu nativa: usa BarraTitulo customizada (minimizar+fechar),
janela sem moldura (FramelessWindowHint), no padrão dos outros projetos.
"""

import os
import subprocess
import sys
from datetime import date
from pathlib import Path

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWidget, QMainWindow, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QComboBox, QPushButton, QFileDialog, QFrame, QScrollArea, QSizePolicy,
)

import config
from frases import frases_da_materia
from caminhos_saida import gerar_caminho_saida
from ui.barra_titulo import BarraTitulo
from ui.dropzone import DropZone
from ui.preview import PreviewPagina
from ui.worker import WorkerFormatacao
from ui.styles import repolir


class JanelaPrincipal(QMainWindow):
    LARGURA = 960
    ALTURA = 760

    def __init__(self):
        super().__init__()
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint)
        self.resize(self.LARGURA, self.ALTURA)

        self._worker = None
        self._caminho_modelo_cabecalho = (
            str(config.CAMINHO_MODELO_CABECALHO_PADRAO)
            if config.CAMINHO_MODELO_CABECALHO_PADRAO.exists()
            else None
        )

        central = QWidget()
        self.setCentralWidget(central)
        layout_geral = QVBoxLayout(central)
        layout_geral.setContentsMargins(0, 0, 0, 0)
        layout_geral.setSpacing(0)

        layout_geral.addWidget(BarraTitulo(self, "FORMATADOR DE PROVAS"))

        corpo = QWidget()
        layout_geral.addWidget(corpo, stretch=1)
        layout_corpo = QHBoxLayout(corpo)
        layout_corpo.setContentsMargins(20, 20, 20, 20)
        layout_corpo.setSpacing(20)

        layout_corpo.addWidget(self._montar_coluna_formulario(), stretch=3)
        layout_corpo.addWidget(self._montar_coluna_preview(), stretch=2)

    # ------------------------------------------------------------------
    # Coluna esquerda: formulário
    # ------------------------------------------------------------------

    def _montar_coluna_formulario(self):
        area_scroll = QScrollArea()
        area_scroll.setWidgetResizable(True)
        area_scroll.setFrameShape(QFrame.Shape.NoFrame)

        conteudo = QWidget()
        area_scroll.setWidget(conteudo)

        layout = QVBoxLayout(conteudo)
        layout.setContentsMargins(0, 0, 10, 0)
        layout.setSpacing(14)

        layout.addWidget(self._rotulo_secao("ARQUIVO DA PROVA"))
        self.dropzone = DropZone()
        self.dropzone.arquivo_selecionado.connect(self._ao_selecionar_arquivo)
        layout.addWidget(self.dropzone)

        layout.addWidget(self._rotulo_secao("DADOS DA PROVA"))

        linha_materia_bimestre = QHBoxLayout()
        linha_materia_bimestre.setSpacing(10)
        linha_materia_bimestre.addWidget(
            self._campo_combo("Matéria", "combo_materia", config.MATERIAS), stretch=1
        )
        linha_materia_bimestre.addWidget(
            self._campo_combo(
                "Bimestre", "combo_bimestre",
                [b["extenso"] for b in config.BIMESTRES],
            ),
            stretch=1,
        )
        layout.addLayout(linha_materia_bimestre)

        linha_professor_turma = QHBoxLayout()
        linha_professor_turma.setSpacing(10)
        linha_professor_turma.addWidget(
            self._campo_texto("Professor(a)", "campo_professor", "Ex.: AMANDA ARAÚJO"),
            stretch=1,
        )
        linha_professor_turma.addWidget(
            self._campo_texto("Série/Turma", "campo_turma", 'Ex.: 2 ANO "A"'),
            stretch=1,
        )
        layout.addLayout(linha_professor_turma)

        linha_data_ano = QHBoxLayout()
        linha_data_ano.setSpacing(10)
        linha_data_ano.addWidget(
            self._campo_texto("Data da prova", "campo_data", "Ex.: 15/09/2026"),
            stretch=1,
        )
        linha_data_ano.addWidget(
            self._campo_texto("Ano letivo (pasta)", "campo_ano", "Ex.: 2026"),
            stretch=1,
        )
        layout.addLayout(linha_data_ano)
        self.campo_ano.setText(str(date.today().year))

        layout.addWidget(
            self._campo_combo("Frase motivacional", "combo_frase", [])
        )

        layout.addWidget(self._rotulo_secao("SAÍDA"))
        container_nome_arquivo = self._campo_texto(
            "Nome do arquivo de saída (sem extensão)",
            "campo_nome_arquivo",
            "Ex.: PROVA_PORTUGUES_2ANO_A",
        )
        layout.addWidget(container_nome_arquivo)

        self.rotulo_destino = QLabel("")
        self.rotulo_destino.setProperty("cls", "status")
        self.rotulo_destino.setWordWrap(True)
        layout.addWidget(self.rotulo_destino)

        self.botao_formatar = QPushButton("FORMATAR PROVA")
        self.botao_formatar.setProperty("variant", "primary")
        self.botao_formatar.setMinimumHeight(40)
        self.botao_formatar.setCursor(Qt.CursorShape.PointingHandCursor)
        self.botao_formatar.clicked.connect(self._ao_clicar_formatar)
        layout.addWidget(self.botao_formatar)

        self.rotulo_status = QLabel("")
        self.rotulo_status.setWordWrap(True)
        layout.addWidget(self.rotulo_status)

        layout.addStretch(1)

        # conecta troca de matéria -> repovoar dropdown de frases
        self.combo_materia.currentTextChanged.connect(self._atualizar_frases_da_materia)
        self.combo_materia.currentTextChanged.connect(self._atualizar_previa_destino)
        self.combo_bimestre.currentTextChanged.connect(self._atualizar_previa_destino)
        self.campo_ano.textChanged.connect(self._atualizar_previa_destino)
        self.campo_nome_arquivo.textChanged.connect(self._atualizar_previa_destino)

        self._atualizar_frases_da_materia(self.combo_materia.currentText())
        self._atualizar_previa_destino()

        return area_scroll

    # ------------------------------------------------------------------
    # Coluna direita: preview + status
    # ------------------------------------------------------------------

    def _montar_coluna_preview(self):
        painel = QWidget()
        painel.setProperty("cls", "painel")
        layout = QVBoxLayout(painel)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        layout.addWidget(self._rotulo_secao("PRÉ-VISUALIZAÇÃO"))

        self.preview = PreviewPagina()
        layout.addWidget(self.preview, alignment=Qt.AlignmentFlag.AlignHCenter)

        layout.addStretch(1)

        self.botao_abrir_pasta = QPushButton("ABRIR PASTA DO ARQUIVO")
        self.botao_abrir_pasta.setEnabled(False)
        self.botao_abrir_pasta.setCursor(Qt.CursorShape.PointingHandCursor)
        self.botao_abrir_pasta.clicked.connect(self._abrir_pasta_resultado)
        layout.addWidget(self.botao_abrir_pasta)

        return painel

    # ------------------------------------------------------------------
    # Helpers de construção de campos
    # ------------------------------------------------------------------

    def _rotulo_secao(self, texto):
        rotulo = QLabel(texto)
        rotulo.setProperty("cls", "rotuloSecao")
        return rotulo

    def _campo_texto(self, rotulo_texto, nome_atributo, placeholder=""):
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        rotulo = QLabel(rotulo_texto)
        rotulo.setProperty("cls", "rotuloCampo")
        layout.addWidget(rotulo)

        campo = QLineEdit()
        campo.setPlaceholderText(placeholder)
        layout.addWidget(campo)

        setattr(self, nome_atributo, campo)
        return container

    def _campo_combo(self, rotulo_texto, nome_atributo, itens):
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        rotulo = QLabel(rotulo_texto)
        rotulo.setProperty("cls", "rotuloCampo")
        layout.addWidget(rotulo)

        combo = QComboBox()
        combo.addItems(itens)
        layout.addWidget(combo)

        setattr(self, nome_atributo, combo)
        return container

    # ------------------------------------------------------------------
    # Comportamento
    # ------------------------------------------------------------------

    def _ao_selecionar_arquivo(self, caminho):
        nome_sugerido = Path(caminho).stem
        if not self.campo_nome_arquivo.text().strip():
            self.campo_nome_arquivo.setText(nome_sugerido)
        self._limpar_status()

    def _atualizar_frases_da_materia(self, materia):
        lista = frases_da_materia(config.CAMINHO_FRASES_JSON, materia)
        self.combo_frase.blockSignals(True)
        self.combo_frase.clear()
        self.combo_frase.addItems(lista)
        self.combo_frase.blockSignals(False)

    def _atualizar_previa_destino(self, *_args):
        ano = self.campo_ano.text().strip() or "----"
        bimestre = self.combo_bimestre.currentText() or "----"
        materia = self.combo_materia.currentText() or "----"
        nome = self.campo_nome_arquivo.text().strip() or "PROVA_FORMATADA"
        self.rotulo_destino.setText(
            f"Será salvo em: provas/{ano}/{bimestre}/{materia}/{nome}.docx"
        )

    def _limpar_status(self):
        self.rotulo_status.setText("")
        self.rotulo_status.setProperty("cls", "status")
        repolir(self.rotulo_status)

    def _mostrar_status_erro(self, mensagem):
        self.rotulo_status.setText(mensagem)
        self.rotulo_status.setProperty("cls", "statusErro")
        repolir(self.rotulo_status)

    def _mostrar_status_sucesso(self, mensagem):
        self.rotulo_status.setText(mensagem)
        self.rotulo_status.setProperty("cls", "statusSucesso")
        repolir(self.rotulo_status)

    def _ao_clicar_formatar(self):
        caminho_entrada = self.dropzone.caminho_atual()
        if not caminho_entrada:
            self._mostrar_status_erro("Selecione o arquivo .docx da prova primeiro.")
            return

        materia = self.combo_materia.currentText()
        bimestre_info = config.BIMESTRES[self.combo_bimestre.currentIndex()]
        ano = self.campo_ano.text().strip() or str(date.today().year)
        nome_arquivo = self.campo_nome_arquivo.text().strip() or Path(caminho_entrada).stem

        caminho_saida = gerar_caminho_saida(
            config.PASTA_PROVAS, ano, bimestre_info["extenso"], materia, nome_arquivo
        )

        valores_cabecalho = {"disciplina": materia}
        if self.campo_professor.text().strip():
            valores_cabecalho["professor"] = self.campo_professor.text().strip()
        if self.campo_turma.text().strip():
            valores_cabecalho["turma"] = self.campo_turma.text().strip()
        if self.campo_data.text().strip():
            valores_cabecalho["data_prova"] = self.campo_data.text().strip()
        if self.combo_frase.currentText().strip():
            valores_cabecalho["frase_motivacional"] = self.combo_frase.currentText().strip()
        valores_cabecalho["bimestre_extenso"] = bimestre_info["extenso"]

        dados = {
            "entrada": caminho_entrada,
            "saida": str(caminho_saida),
            "disciplina": materia,
            "bimestre": bimestre_info["curto"],
            "modelo_cabecalho": self._caminho_modelo_cabecalho,
            "valores_cabecalho": valores_cabecalho,
        }

        self._ultimo_caminho_saida = str(caminho_saida)
        self._definir_carregando(True)
        self.preview.limpar()

        self._worker = WorkerFormatacao(dados)
        self._worker.sucesso.connect(self._ao_formatar_com_sucesso)
        self._worker.erro.connect(self._ao_formatar_com_erro)
        self._worker.start()

    def _definir_carregando(self, carregando):
        self.botao_formatar.setEnabled(not carregando)
        self.botao_formatar.setText("FORMATANDO..." if carregando else "FORMATAR PROVA")
        if carregando:
            self._limpar_status()

    def _ao_formatar_com_sucesso(self, caminho_docx, caminho_pdf):
        self._definir_carregando(False)
        self._mostrar_status_sucesso(f"Prova formatada com sucesso:\n{Path(caminho_docx).name}")
        self.botao_abrir_pasta.setEnabled(True)
        if caminho_pdf:
            self.preview.mostrar_pdf(caminho_pdf)
        else:
            self.preview.limpar()

    def _ao_formatar_com_erro(self, mensagem):
        self._definir_carregando(False)
        self._mostrar_status_erro(mensagem)

    def _abrir_pasta_resultado(self):
        if not getattr(self, "_ultimo_caminho_saida", None):
            return
        pasta = str(Path(self._ultimo_caminho_saida).parent)
        if sys.platform.startswith("win"):
            os.startfile(pasta)  # noqa: S606 - abrir explorador é o comportamento desejado
        elif sys.platform == "darwin":
            subprocess.run(["open", pasta])
        else:
            subprocess.run(["xdg-open", pasta])
