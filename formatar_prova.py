#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
formatar_prova.py
------------------
Reformata uma prova .docx "crua" (como sai do professor) para o padrão
visual da escola usado nas provas já diagramadas:

  - Caixa de título laranja clara: "ATIVIDADE AVALIATIVA DE [DISCIPLINA] – 3º BIMESTRE"
  - Fonte Calibri em todo o corpo do documento
  - Título da caixa em 14pt / resto do texto em 11pt
  - Layout em 2 colunas
  - Linha vertical preta dividindo a página ao meio
  - Espaçamento e recuo padronizados
  - Quebra de texto das imagens padronizada (inline, alinhada ao texto)
  - Imagens redimensionadas para caber dentro da coluna

USO:
    python formatar_prova.py entrada.docx saida.docx --disciplina "PORTUGUÊS"

O script tenta detectar a disciplina automaticamente a partir do título
existente no documento ("AVALIAÇÃO DE LÍNGUA PORTUGUESA...") caso você não
informe --disciplina.

Limitações conhecidas (o que o script NÃO faz):
  - Não recria o cabeçalho decorado (logo da escola, campos Nome/Nº/Turma/
    Professor/Data/Frase motivacional) porque esse cabeçalho é uma peça de
    identidade visual fixa da escola, não algo derivável do conteúdo da
    prova crua. Ele pode ser colado de um modelo existente (ver função
    `inserir_cabecalho_modelo`) se você fornecer um arquivo-modelo.
  - Questões muito grandes (ex.: um poema com imagem alta) podem ficar
    "cortadas" na quebra de coluna — isso é uma limitação do próprio motor
    de colunas do Word/LibreOffice, não do script. Revise visualmente.
"""

import argparse
import copy
import re
import sys
from pathlib import Path

from docx import Document
from docx.shared import Pt, Cm, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image


# --------------------------------------------------------------------------
# Configurações de estilo (extraídas do modelo de prova da escola)
# --------------------------------------------------------------------------

FONTE_CORPO = "Calibri"
FONTE_TITULO = "Montserrat"          # fonte usada dentro da caixa de título
TAMANHO_TITULO_PT = 14
TAMANHO_CORPO_PT = 11

COR_CAIXA_TITULO = "FDEADA"          # laranja bem claro (fundo da caixa)
COR_LINHA_DIVISORIA = "000000"       # preto
LARGURA_LINHA_DIVISORIA_PT = 1.0     # ~12700 EMU

# Margens de página (iguais ao modelo: 567 twips ≈ 1 cm)
MARGEM_CM = 1.0

# Recuo padrão de parágrafo (0 = sem recuo, igual ao modelo)
RECUO_PADRAO_CM = 0.0
ESPACO_DEPOIS_PARAGRAFO_PT = 6

# Largura máxima de imagem dentro de uma coluna (em cm), com folga para
# não estourar a margem interna da coluna de ~2 colunas numa A4 com
# margens de 1 cm de cada lado.
LARGURA_MAX_IMAGEM_COLUNA_CM = 7.8


# --------------------------------------------------------------------------
# Utilitários de baixo nível (OOXML)
# --------------------------------------------------------------------------

def definir_fonte_padrao_documento(doc, fonte=FONTE_CORPO, tamanho_pt=TAMANHO_CORPO_PT):
    """Ajusta o estilo 'Normal' para usar a fonte/tamanho padrão do corpo."""
    estilo_normal = doc.styles["Normal"]
    estilo_normal.font.name = fonte
    estilo_normal.font.size = Pt(tamanho_pt)

    # rFonts precisa também dos atributos eastAsia/cs para valer em todo
    # lugar (Word decide a fonte a partir de w:eastAsia/w:cs dependendo do
    # idioma do texto, então setar só "name" as vezes não é suficiente).
    rpr = estilo_normal.element.get_or_add_rPr()
    rFonts = rpr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rpr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), fonte)


def forcar_fonte_em_todos_os_runs(doc, fonte=FONTE_CORPO, tamanho_pt=TAMANHO_CORPO_PT):
    """
    Percorre TODOS os runs do documento (corpo + dentro de textboxes) e
    força fonte/tamanho. Isso é necessário porque a formatação direta em
    cada run (w:rFonts no rPr do próprio run) tem prioridade sobre o
    estilo 'Normal' do documento.
    """
    # Runs "normais" (parágrafos do corpo, incluindo os de dentro de tabelas)
    for paragrafo in iterar_todos_os_paragrafos(doc):
        for run in paragrafo.runs:
            _aplicar_fonte_no_run(run, fonte, tamanho_pt)

    # Runs dentro de textboxes / shapes (não aparecem via doc.paragraphs)
    corpo_xml = doc.element.body
    for rpr in corpo_xml.iter(qn("w:rPr")):
        _forcar_fonte_no_rpr(rpr, fonte)


def _aplicar_fonte_no_run(run, fonte, tamanho_pt):
    run.font.name = fonte
    run.font.size = Pt(tamanho_pt)
    rpr = run._element.get_or_add_rPr()
    _forcar_fonte_no_rpr(rpr, fonte)


def _forcar_fonte_no_rpr(rpr, fonte):
    rFonts = rpr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rpr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), fonte)


def iterar_todos_os_paragrafos(doc):
    """Gera todos os parágrafos do corpo do documento, incluindo os que
    estão dentro de tabelas (mas não dentro de textboxes — python-docx
    não enxerga o conteúdo de textboxes como Paragraph nativamente)."""
    for paragrafo in doc.paragraphs:
        yield paragrafo
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for paragrafo in celula.paragraphs:
                    yield paragrafo


def padronizar_paragrafos(doc):
    """Padroniza recuo e espaçamento de todos os parágrafos do corpo."""
    for paragrafo in iterar_todos_os_paragrafos(doc):
        pf = paragrafo.paragraph_format
        pf.left_indent = Cm(RECUO_PADRAO_CM)
        pf.first_line_indent = None
        pf.space_after = Pt(ESPACO_DEPOIS_PARAGRAFO_PT)
        pf.space_before = Pt(0)


def configurar_pagina_duas_colunas(section, num_colunas=2, espaco_entre_colunas_cm=0.6):
    """Configura a seção para A4, margens de 1 cm e duas colunas de texto."""
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(MARGEM_CM)
    section.right_margin = Cm(MARGEM_CM)
    section.top_margin = Cm(MARGEM_CM)
    section.bottom_margin = Cm(MARGEM_CM)

    sectPr = section._sectPr
    # remove w:cols existente, se houver
    for cols in sectPr.findall(qn("w:cols")):
        sectPr.remove(cols)
    cols = OxmlElement("w:cols")
    cols.set(qn("w:num"), str(num_colunas))
    cols.set(qn("w:space"), str(int(espaco_entre_colunas_cm * 567)))  # cm -> twips aprox (1cm=567twips)
    sectPr.append(cols)


# --------------------------------------------------------------------------
# Cabeçalho decorado da escola (logo + campos), copiado de um .docx-modelo
# --------------------------------------------------------------------------

# Nome das formas (docPr) dentro do primeiro parágrafo do modelo que compõem
# o cabeçalho decorado da escola, e o que cada uma representa. A "Caixa de
# Texto 2" (título) é deliberadamente excluída daqui porque já é recriada
# dinamicamente por inserir_caixa_titulo().
CAMPOS_CABECALHO_MODELO = {
    "Caixa de Texto 3": "professor",
    "Caixa de Texto 5": "disciplina",
    "Caixa de Texto 6": "data_prova",
    "Caixa de Texto 7": "turma",
    "Caixa de Texto 9": "frase_motivacional",
    "Caixa de Texto 14": "bimestre_extenso",
    # "Caixa de Texto 4" = linha para nome do aluno (não costuma mudar)
    # "Imagem 1" = logo da escola + moldura (não costuma mudar)
}


def inserir_cabecalho_modelo(doc, caminho_docx_modelo, valores=None):
    """
    Copia o cabeçalho decorado da escola (logo da escola + moldura + campos
    Nome do Aluno / Nº / Série-Turma / Professor(a) / Disciplina /
    Data da Prova / Bimestre / Frase Motivacional) de um .docx-modelo já
    diagramado, e insere no topo do documento `doc`.

    Isso resolve a parte que NÃO dá pra derivar do conteúdo da prova crua:
    a identidade visual fixa da escola. O modelo é reaproveitado tal como
    está — só o texto dos campos é opcionalmente trocado via `valores`.

    Parâmetros
    ----------
    doc : Document
        Documento de destino (já com a caixa de título e o resto formatado).
    caminho_docx_modelo : str
        Caminho para um .docx que contenha o cabeçalho decorado no primeiro
        parágrafo (ex.: qualquer prova já diagramada no padrão da escola).
    valores : dict, opcional
        Sobrescreve o texto de campos específicos do cabeçalho. Chaves
        aceitas: "professor", "disciplina", "data_prova", "turma",
        "frase_motivacional", "bimestre_extenso". Campos omitidos mantêm o
        texto exatamente como está no modelo.

    Retorna
    -------
    bool
        True se o cabeçalho foi encontrado e inserido; False se o modelo
        informado não tinha um cabeçalho reconhecível (nesse caso o
        documento de destino não é alterado).
    """
    valores = valores or {}
    modelo = Document(caminho_docx_modelo)
    corpo_modelo = modelo.element.body
    primeiro_paragrafo_modelo = corpo_modelo.find(qn("w:p"))
    if primeiro_paragrafo_modelo is None:
        return False

    formas_encontradas = []
    for anchor in primeiro_paragrafo_modelo.iter(qn("wp:anchor")):
        docpr = anchor.find(qn("wp:docPr"))
        nome = docpr.get("name") if docpr is not None else None
        if nome == "Caixa de Texto 2":
            continue  # é o título — não faz parte do "cabeçalho fixo"
        formas_encontradas.append((nome, anchor))

    if not formas_encontradas:
        return False

    # cria o parágrafo de destino que vai hospedar as formas copiadas
    novo_paragrafo = OxmlElement("w:p")

    for nome, anchor_original in formas_encontradas:
        anchor_copiado = copy.deepcopy(anchor_original)

        campo = CAMPOS_CABECALHO_MODELO.get(nome)
        if campo and campo in valores:
            _substituir_texto_da_forma(anchor_copiado, valores[campo])

        # imagens (ex.: o logo) precisam da relação (rId) recriada no
        # documento de destino, senão a imagem quebra
        _remapear_imagens_da_forma(anchor_copiado, modelo, doc)

        run_novo = OxmlElement("w:r")
        drawing_novo = OxmlElement("w:drawing")
        drawing_novo.append(anchor_copiado)
        run_novo.append(drawing_novo)
        novo_paragrafo.append(run_novo)

    doc.element.body.insert(0, novo_paragrafo)
    return True


def _substituir_texto_da_forma(anchor, novo_texto):
    """Troca o texto de dentro de wps:txbx/w:txbxContent por `novo_texto`,
    mantendo a formatação do PRIMEIRO run de texto encontrado (todos os
    demais runs de texto são removidos para não duplicar o conteúdo)."""
    txbx_content = anchor.find(".//" + qn("w:txbxContent"))
    if txbx_content is None:
        return
    primeiro_paragrafo_txbx = txbx_content.find(qn("w:p"))
    if primeiro_paragrafo_txbx is None:
        return
    runs = primeiro_paragrafo_txbx.findall(qn("w:r"))
    if not runs:
        return
    primeiro_run = runs[0]
    for run_extra in runs[1:]:
        primeiro_paragrafo_txbx.remove(run_extra)
    elemento_t = primeiro_run.find(qn("w:t"))
    if elemento_t is None:
        elemento_t = OxmlElement("w:t")
        primeiro_run.append(elemento_t)
    elemento_t.set(qn("xml:space"), "preserve")
    elemento_t.text = novo_texto

    # remove parágrafos extras de texto dentro da mesma textbox (caso a
    # caixa original tivesse mais de uma linha, ex.: quebras manuais)
    for paragrafo_extra in txbx_content.findall(qn("w:p"))[1:]:
        txbx_content.remove(paragrafo_extra)


def _remapear_imagens_da_forma(anchor, doc_origem, doc_destino):
    """Para cada <a:blip r:embed="rIdX"> dentro da forma copiada, importa a
    imagem correspondente do documento de origem para o documento de
    destino e atualiza o r:embed com o novo rId gerado."""
    for blip in anchor.iter(qn("a:blip")):
        rid_antigo = blip.get(qn("r:embed"))
        if not rid_antigo:
            continue
        parte_imagem = doc_origem.part.related_parts.get(rid_antigo)
        if parte_imagem is None:
            continue
        bytes_imagem = parte_imagem.blob
        import io
        novo_rid, _ = doc_destino.part.get_or_add_image(io.BytesIO(bytes_imagem))
        blip.set(qn("r:embed"), novo_rid)


# --------------------------------------------------------------------------
# Caixa de título (textbox laranja)
# --------------------------------------------------------------------------

def inserir_caixa_titulo(doc, disciplina, bimestre="3º"):
    """
    Insere, como PRIMEIRO parágrafo do documento, uma caixa de texto
    retangular com fundo laranja claro contendo:
        "ATIVIDADE AVALIATIVA DE [DISCIPLINA] – [N]º BIMESTRE"
    reproduzindo o textbox usado no modelo da escola.
    """
    texto_titulo = f"ATIVIDADE AVALIATIVA DE {disciplina.upper()} – {bimestre} BIMESTRE"

    corpo = doc.element.body
    novo_paragrafo_xml = _construir_xml_caixa_titulo(texto_titulo)
    corpo.insert(0, novo_paragrafo_xml)

    # parágrafo vazio logo abaixo para dar respiro antes do conteúdo
    p_vazio = OxmlElement("w:p")
    corpo.insert(1, p_vazio)


def _construir_xml_caixa_titulo(texto):
    """Monta o XML de um parágrafo contendo uma textbox (wps:wsp) com o
    título, no mesmo padrão visual da caixa laranja do modelo da escola."""

    largura_emu = 5267325   # ~ largura útil de uma A4 com margens de 1cm
    altura_emu = 333375

    xml = f"""
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
     xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
     xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:pPr><w:jc w:val="center"/></w:pPr>
  <w:r>
    <w:rPr><w:rFonts w:ascii="{FONTE_CORPO}" w:hAnsi="{FONTE_CORPO}"/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="45720" distB="45720" distL="114300" distR="114300"
                     simplePos="0" relativeHeight="251667456" behindDoc="0"
                     locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="{largura_emu}" cy="{altura_emu}"/>
            <wp:effectExtent l="0" t="0" r="9525" b="9525"/>
            <wp:wrapSquare wrapText="bothSides"/>
            <wp:docPr id="9001" name="Caixa de Titulo Prova"/>
            <wp:cNvGraphicFramePr/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr bwMode="auto">
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="{largura_emu}" cy="{altura_emu}"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="{COR_CAIXA_TITULO}"/></a:solidFill>
                    <a:ln w="9525"><a:noFill/></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p>
                        <w:pPr>
                          <w:spacing w:after="0"/>
                          <w:jc w:val="center"/>
                        </w:pPr>
                        <w:r>
                          <w:rPr>
                            <w:rFonts w:ascii="{FONTE_TITULO}" w:hAnsi="{FONTE_TITULO}"/>
                            <w:b/>
                            <w:sz w:val="{TAMANHO_TITULO_PT * 2}"/>
                            <w:szCs w:val="{TAMANHO_TITULO_PT * 2}"/>
                          </w:rPr>
                          <w:t xml:space="preserve">{_escapar_xml(texto)}</w:t>
                        </w:r>
                      </w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440"
                              tIns="45720" rIns="91440" bIns="45720"
                              anchor="ctr" anchorCtr="0">
                    <a:noAutofit/>
                  </wps:bodyPr>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
      <mc:Fallback>
        <w:pict/>
      </mc:Fallback>
    </mc:AlternateContent>
  </w:r>
</w:p>
"""
    from lxml import etree
    return etree.fromstring(xml.encode("utf-8"))


def _escapar_xml(texto):
    return (
        texto.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


# --------------------------------------------------------------------------
# Linha vertical dividindo as colunas
# --------------------------------------------------------------------------

def inserir_linha_divisoria_central(doc):
    """
    Insere um 'conector reto' vertical preto ancorado no meio da página,
    cobrindo do topo ao fundo da área útil — reproduzindo a linha que
    separa visualmente as duas colunas no modelo da escola.

    A posição horizontal é calculada a partir da largura da página menos
    as margens (ou seja, o meio exato da área útil).
    """
    section = doc.sections[0]
    largura_pagina_emu = section.page_width
    margem_esq_emu = section.left_margin
    margem_dir_emu = section.right_margin
    largura_util_emu = largura_pagina_emu - margem_esq_emu - margem_dir_emu
    posicao_x_emu = int(margem_esq_emu + largura_util_emu / 2)

    altura_pagina_emu = section.page_height
    margem_sup_emu = section.top_margin
    margem_inf_emu = section.bottom_margin
    altura_linha_emu = int(altura_pagina_emu - margem_sup_emu - margem_inf_emu)

    largura_linha_emu = Pt(LARGURA_LINHA_DIVISORIA_PT).emu

    xml = f"""
<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
     xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
     xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:pPr><w:spacing w:after="0"/></w:pPr>
  <w:r>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300"
                     simplePos="0" relativeHeight="251685888" behindDoc="0"
                     locked="0" layoutInCell="0" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="page"><wp:posOffset>{posicao_x_emu}</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="page"><wp:posOffset>{margem_sup_emu}</wp:posOffset></wp:positionV>
            <wp:extent cx="{largura_linha_emu}" cy="{altura_linha_emu}"/>
            <wp:effectExtent l="6350" t="0" r="19050" b="6985"/>
            <wp:wrapNone/>
            <wp:docPr id="9002" name="Linha Divisoria Central"/>
            <wp:cNvGraphicFramePr/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvCnPr/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="{largura_linha_emu}" cy="{altura_linha_emu}"/></a:xfrm>
                    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
                    <a:ln w="{largura_linha_emu}">
                      <a:solidFill><a:srgbClr val="{COR_LINHA_DIVISORIA}"/></a:solidFill>
                    </a:ln>
                  </wps:spPr>
                  <wps:bodyPr/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
      <mc:Fallback><w:pict/></mc:Fallback>
    </mc:AlternateContent>
  </w:r>
</w:p>
"""
    from lxml import etree
    elemento = etree.fromstring(xml.encode("utf-8"))
    doc.element.body.append(elemento)


# --------------------------------------------------------------------------
# Redimensionamento de imagens
# --------------------------------------------------------------------------

def converter_imagens_ancoradas_em_inline(doc):
    """
    Converte imagens que estão 'flutuando' (wp:anchor, posição absoluta na
    página) para o modo 'inline' (fluem junto com o texto, como um
    caractere gigante no meio do parágrafo).

    Por quê: imagens flutuantes não se movem quando o texto se re-flui em
    colunas — elas ficam grudadas numa posição fixa da página, o que faz
    a imagem de uma questão "descolar" do enunciado dela quando o layout
    muda para 2 colunas (a imagem pode acabar entrando na coluna errada).
    Deixando inline, a imagem sempre acompanha o parágrafo em que foi
    inserida, não importa como as colunas se ajustem — isso padroniza a
    quebra de texto de todas as imagens da prova.

    Ignora âncoras que não são imagens de verdade (textboxes, linhas,
    formas), identificadas pela ausência de <pic:pic> dentro delas.
    """
    corpo = doc.element.body
    anchors = [a for a in corpo.iter(qn("wp:anchor")) if a.find(".//" + qn("pic:pic")) is not None]

    for anchor in anchors:
        extent = anchor.find(qn("wp:extent"))
        doc_pr = anchor.find(qn("wp:docPr"))
        graphic = anchor.find(qn("a:graphic"))
        if extent is None or graphic is None:
            continue

        cx = extent.get("cx")
        cy = extent.get("cy")

        # cria um novo wp:inline com o mesmo conteúdo gráfico
        inline = OxmlElement("wp:inline")
        inline.set("distT", "0")
        inline.set("distB", "0")
        inline.set("distL", "0")
        inline.set("distR", "0")

        novo_extent = OxmlElement("wp:extent")
        novo_extent.set("cx", cx)
        novo_extent.set("cy", cy)
        inline.append(novo_extent)

        effect_extent_original = anchor.find(qn("wp:effectExtent"))
        if effect_extent_original is not None:
            inline.append(copy.deepcopy(effect_extent_original))

        if doc_pr is not None:
            inline.append(copy.deepcopy(doc_pr))
        else:
            fallback_doc_pr = OxmlElement("wp:docPr")
            fallback_doc_pr.set("id", "1")
            fallback_doc_pr.set("name", "Imagem")
            inline.append(fallback_doc_pr)

        cnv_graphic_frame_pr = anchor.find(qn("wp:cNvGraphicFramePr"))
        if cnv_graphic_frame_pr is not None:
            inline.append(copy.deepcopy(cnv_graphic_frame_pr))

        inline.append(copy.deepcopy(graphic))

        # troca a wp:anchor pela wp:inline dentro do w:drawing pai
        drawing = anchor.getparent()
        drawing.replace(anchor, inline)


def separar_imagens_de_texto_na_mesma_linha(doc):
    """
    Quando um parágrafo mistura, no mesmo w:p, um run com imagem (w:drawing)
    e um run de texto (comum em provas onde a imagem foi colada 'no meio'
    da frase por acidente), move a imagem para um parágrafo próprio,
    logo ANTES do parágrafo de texto. Isso evita que a imagem inline
    empurre a primeira linha do enunciado para o lado, deixando o texto
    corrido e alinhado como as demais questões.
    """
    corpo = doc.element.body
    paragrafos = list(corpo.findall(qn("w:p")))

    for paragrafo in paragrafos:
        runs_com_imagem = [
            r for r in paragrafo.findall(qn("w:r"))
            if r.find(".//" + qn("w:drawing")) is not None
        ]
        runs_com_texto = [
            r for r in paragrafo.findall(qn("w:r"))
            if r.find(qn("w:t")) is not None and (r.find(qn("w:t")).text or "").strip()
        ]
        if not runs_com_imagem or not runs_com_texto:
            continue  # não é um parágrafo "misto"

        # cria um novo parágrafo só com a(s) imagem(ns), inserido antes
        novo_paragrafo = OxmlElement("w:p")
        pPr_original = paragrafo.find(qn("w:pPr"))
        if pPr_original is not None:
            novo_paragrafo.append(copy.deepcopy(pPr_original))
        for run_imagem in runs_com_imagem:
            novo_paragrafo.append(copy.deepcopy(run_imagem))
            paragrafo.remove(run_imagem)

        paragrafo.addprevious(novo_paragrafo)


def redimensionar_imagens(doc, largura_max_cm=LARGURA_MAX_IMAGEM_COLUNA_CM):
    """
    Percorre todas as imagens (inline e ancoradas) do documento e, quando a
    largura ultrapassa `largura_max_cm`, reduz proporcionalmente para caber
    dentro da largura de uma coluna, preservando a proporção original.
    Também padroniza a quebra de texto das imagens para 'quadrado'
    (wrapSquare) quando a imagem está ancorada (flutuante), para ficar
    igual ao padrão do modelo.
    """
    largura_max_emu = Cm(largura_max_cm).emu

    # --- imagens inline (w:drawing/wp:inline) ---
    for inline in doc.element.body.iter(qn("wp:inline")):
        extent = inline.find(qn("wp:extent"))
        if extent is None:
            continue
        cx = int(extent.get("cx"))
        cy = int(extent.get("cy"))
        if cx > largura_max_emu:
            fator = largura_max_emu / cx
            extent.set("cx", str(int(cx * fator)))
            extent.set("cy", str(int(cy * fator)))
            _sincronizar_xfrm_da_imagem(inline, int(cx * fator), int(cy * fator))

    # --- imagens ancoradas (w:drawing/wp:anchor) que sejam <pic:pic> de fato
    # (ignora as textboxes/linhas que também usam wp:anchor) ---
    for anchor in doc.element.body.iter(qn("wp:anchor")):
        if anchor.find(".//" + qn("pic:pic")) is None:
            continue  # é textbox/linha, não imagem
        extent = anchor.find(qn("wp:extent"))
        if extent is None:
            continue
        cx = int(extent.get("cx"))
        cy = int(extent.get("cy"))
        if cx > largura_max_emu:
            fator = largura_max_emu / cx
            extent.set("cx", str(int(cx * fator)))
            extent.set("cy", str(int(cy * fator)))
            _sincronizar_xfrm_da_imagem(anchor, int(cx * fator), int(cy * fator))

        # padroniza quebra de texto -> quadrado (comportamento do modelo)
        for tag in ("wrapNone", "wrapTight", "wrapThrough", "wrapTopAndBottom"):
            existente = anchor.find(qn(f"wp:{tag}"))
            if existente is not None:
                anchor.remove(existente)
        if anchor.find(qn("wp:wrapSquare")) is None:
            wrap_square = OxmlElement("wp:wrapSquare")
            wrap_square.set("wrapText", "bothSides")
            # precisa ficar na posição correta do schema (depois de effectExtent)
            effect_extent = anchor.find(qn("wp:effectExtent"))
            if effect_extent is not None:
                effect_extent.addnext(wrap_square)
            else:
                anchor.insert(3, wrap_square)


def _sincronizar_xfrm_da_imagem(drawing_container, novo_cx, novo_cy):
    """Atualiza também a:xfrm/a:ext dentro do pic:spPr para bater com o
    novo tamanho definido em wp:extent (evita imagem 'esticada')."""
    for ext in drawing_container.iter(qn("a:ext")):
        pai = ext.getparent()
        if pai is not None and pai.tag == qn("a:xfrm"):
            ext.set("cx", str(novo_cx))
            ext.set("cy", str(novo_cy))


# --------------------------------------------------------------------------
# Detecção de disciplina a partir do texto existente
# --------------------------------------------------------------------------

def detectar_disciplina(doc, fallback="PORTUGUÊS"):
    """
    Tenta adivinhar a disciplina a partir do primeiro título encontrado no
    documento (ex.: 'AVALIAÇÃO DE LÍNGUA PORTUGUESA – 2° ANO' -> 'PORTUGUÊS').
    Se não conseguir reconhecer um padrão conhecido, usa o fallback.
    """
    mapa_disciplinas = {
        r"portugu": "PORTUGUÊS",
        r"matem[aá]tica": "MATEMÁTICA",
        r"ci[eê]ncias": "CIÊNCIAS",
        r"hist[oó]ria": "HISTÓRIA",
        r"geografia": "GEOGRAFIA",
        r"ingl[eê]s": "INGLÊS",
        r"artes?": "ARTES",
        r"ensino religioso": "ENSINO RELIGIOSO",
        r"educa[cç][aã]o f[ií]sica": "EDUCAÇÃO FÍSICA",
    }
    for paragrafo in doc.paragraphs[:15]:
        texto = paragrafo.text.strip()
        if not texto:
            continue
        texto_normalizado = texto.lower()
        for padrao, nome in mapa_disciplinas.items():
            if re.search(padrao, texto_normalizado):
                return nome
    return fallback


def remover_titulo_antigo(doc, disciplina_detectada_texto=None):
    """
    Remove o(s) parágrafo(s) de título antigo do tipo
    'AVALIAÇÃO DE LÍNGUA PORTUGUESA – 2° ANO' para não duplicar com a nova
    caixa de título. Procura por padrões comuns de título de prova nos
    primeiros parágrafos.
    """
    padroes_titulo = [
        r"avalia[cç][aã]o de",
        r"prova de",
        r"atividade avaliativa",
    ]
    for paragrafo in list(doc.paragraphs[:10]):
        texto = paragrafo.text.strip().lower()
        if not texto:
            continue
        if any(re.search(p, texto) for p in padroes_titulo):
            elemento = paragrafo._element
            elemento.getparent().remove(elemento)


def remover_cabecalho_manual_antigo(doc):
    """
    Remove parágrafos/textboxes de cabeçalho manual do tipo
    'ALUNO: ____', 'ESCOLA: ____ TURMA: ____', 'PROFESSOR: ____ DATA: ____'
    que os professores costumam digitar à mão no topo da prova crua.
    Esses campos batem com os que já existem no cabeçalho oficial da escola,
    então ficam redundantes/colidindo com a caixa de título nova.

    Cobre dois casos:
      1) Texto normal no corpo (w:p com w:t contendo os rótulos).
      2) Texto dentro de textboxes flutuantes (w:txbxContent).
    """
    padroes_cabecalho = re.compile(
        r"^(ALUNO|ESCOLA|PROFESSOR|TURMA|DATA)\s*:", re.IGNORECASE
    )

    # Caso 1: parágrafos "soltos" no corpo do documento
    for paragrafo in list(doc.paragraphs[:20]):
        texto = paragrafo.text.strip()
        if texto and padroes_cabecalho.match(texto):
            elemento = paragrafo._element
            elemento.getparent().remove(elemento)

    # Caso 2: parágrafos dentro de textboxes (w:txbxContent) — localizamos
    # o w:p pai mais próximo (o parágrafo do corpo que ANCORA a textbox)
    # e removemos ele inteiro se o texto do cabeçalho aparecer lá dentro.
    corpo = doc.element.body
    for txbx_content in corpo.iter(qn("w:txbxContent")):
        texto_txbx = "".join(
            t.text or "" for t in txbx_content.iter(qn("w:t"))
        ).strip()
        if texto_txbx and padroes_cabecalho.match(texto_txbx):
            # sobe até achar o w:p do corpo principal que contém essa textbox
            paragrafo_ancora = _encontrar_paragrafo_ancora(txbx_content, corpo)
            if paragrafo_ancora is not None and paragrafo_ancora.getparent() is not None:
                paragrafo_ancora.getparent().remove(paragrafo_ancora)


def _encontrar_paragrafo_ancora(elemento, corpo):
    """Sobe na árvore XML a partir de `elemento` até encontrar o w:p que é
    filho direto de `corpo` (ou seja, o parágrafo do documento principal
    que hospeda essa textbox)."""
    atual = elemento
    while atual is not None:
        pai = atual.getparent()
        if pai is corpo and atual.tag == qn("w:p"):
            return atual
        atual = pai
    return None


# --------------------------------------------------------------------------
# Fluxo principal
# --------------------------------------------------------------------------

def formatar_prova(caminho_entrada, caminho_saida, disciplina=None, bimestre="3º",
                    caminho_docx_modelo_cabecalho=None, valores_cabecalho=None):
    doc = Document(caminho_entrada)

    if disciplina is None:
        disciplina = detectar_disciplina(doc)

    # 1) remove título antigo e cabeçalho manual (evita duplicar/colidir
    #    com a caixa de título nova)
    remover_titulo_antigo(doc)
    remover_cabecalho_manual_antigo(doc)

    # 2) fonte padrão do documento inteiro
    definir_fonte_padrao_documento(doc, FONTE_CORPO, TAMANHO_CORPO_PT)
    forcar_fonte_em_todos_os_runs(doc, FONTE_CORPO, TAMANHO_CORPO_PT)

    # 3) espaçamento / recuo padronizados
    padronizar_paragrafos(doc)

    # 4) padroniza quebra de texto das imagens (todas inline), separa
    #    imagem de texto quando estão coladas na mesma linha, e redimensiona
    converter_imagens_ancoradas_em_inline(doc)
    separar_imagens_de_texto_na_mesma_linha(doc)
    redimensionar_imagens(doc)

    # 5) layout de página em 2 colunas
    section = doc.sections[0]
    configurar_pagina_duas_colunas(section)

    # 6) caixa de título laranja no topo
    inserir_caixa_titulo(doc, disciplina, bimestre)

    # 7) linha vertical dividindo as colunas
    inserir_linha_divisoria_central(doc)

    # 8) (opcional) cabeçalho decorado da escola, copiado de um modelo
    if caminho_docx_modelo_cabecalho:
        inserir_cabecalho_modelo(doc, caminho_docx_modelo_cabecalho, valores_cabecalho)

    doc.save(caminho_saida)
    return disciplina


def main():
    parser = argparse.ArgumentParser(
        description="Reformata uma prova .docx para o padrão visual da escola."
    )
    parser.add_argument("entrada", help="Caminho do .docx original (cru)")
    parser.add_argument("saida", help="Caminho do .docx de saída (formatado)")
    parser.add_argument(
        "--disciplina",
        default=None,
        help="Nome da disciplina para o título (ex.: PORTUGUÊS). "
             "Se omitido, tenta detectar automaticamente.",
    )
    parser.add_argument(
        "--bimestre",
        default="3º",
        help="Texto do bimestre para o título (padrão: '3º')",
    )
    parser.add_argument(
        "--modelo-cabecalho",
        default=None,
        help="Caminho de um .docx já diagramado no padrão da escola de onde "
             "copiar o cabeçalho decorado (logo + campos Nome/Nº/Turma/"
             "Professor/Data/Frase motivacional). Se omitido, o documento "
             "sai sem esse cabeçalho decorado.",
    )
    parser.add_argument(
        "--professor", default=None, help="Sobrescreve o campo Professor(a) do cabeçalho"
    )
    parser.add_argument(
        "--turma", default=None, help="Sobrescreve o campo Série/Turma do cabeçalho"
    )
    parser.add_argument(
        "--data-prova", default=None, help="Sobrescreve o campo Data da Prova do cabeçalho"
    )
    parser.add_argument(
        "--frase-motivacional", default=None, help="Sobrescreve a frase motivacional do cabeçalho"
    )
    parser.add_argument(
        "--bimestre-extenso",
        default=None,
        help="Sobrescreve o campo Bimestre do cabeçalho (ex.: 'III BIMESTRE'). "
             "Diferente de --bimestre, que é só o número usado no título.",
    )
    args = parser.parse_args()

    valores_cabecalho = {}
    if args.professor:
        valores_cabecalho["professor"] = args.professor
    if args.turma:
        valores_cabecalho["turma"] = args.turma
    if args.data_prova:
        valores_cabecalho["data_prova"] = args.data_prova
    if args.frase_motivacional:
        valores_cabecalho["frase_motivacional"] = args.frase_motivacional
    if args.bimestre_extenso:
        valores_cabecalho["bimestre_extenso"] = args.bimestre_extenso
    # a disciplina do cabeçalho acompanha a mesma usada no título, a menos
    # que o usuário sobrescreva manualmente via --disciplina
    if args.disciplina:
        valores_cabecalho["disciplina"] = args.disciplina

    disciplina_usada = formatar_prova(
        args.entrada,
        args.saida,
        disciplina=args.disciplina,
        bimestre=args.bimestre,
        caminho_docx_modelo_cabecalho=args.modelo_cabecalho,
        valores_cabecalho=valores_cabecalho or None,
    )
    print(f"OK. Disciplina usada no título: {disciplina_usada}")
    print(f"Arquivo salvo em: {args.saida}")


if __name__ == "__main__":
    main()
