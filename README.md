# Formatador de Provas (interface gráfica)

Interface PyQt6, na identidade visual da escola (tema escuro + laranja
`#db6018`, barra de título customizada), para o `formatar_prova.py`.
Arraste o `.docx` cru da prova (ou clique para escolher pelo explorador
de arquivos do Windows), preencha os dados, e gere a prova já formatada
no padrão da escola — com preview da primeira página dentro do próprio
app.

## Download

Baixe a versão pronta para Windows na página **Releases** deste repositório
(arquivo `.zip`; descompacte e abra o `.exe`). O Windows pode mostrar um
aviso de segurança na primeira vez, porque o programa não tem assinatura
digital: clique em **Mais informações → Executar assim mesmo**.

## Privacidade

Este repositório tem **apenas o código**. Nenhum dado de aluno, planilha,
banco de dados ou credencial faz parte dele: esses arquivos ficam só no
computador da escola e estão bloqueados pelo `.gitignore`.

## Instalação

```bash
pip install PyQt6 python-docx pillow lxml pymupdf
```

- `pymupdf` é opcional — só é usado para desenhar a pré-visualização da
  primeira página dentro do app. Sem ele, o app formata normalmente, só
  não mostra a miniatura.
- Também é preciso ter o **LibreOffice** instalado no computador (e
  acessível como `soffice` no PATH do Windows), porque é ele quem gera o
  PDF usado no preview. Sem o LibreOffice, a formatação funciona igual —
  só o preview fica em branco.

## Rodar

```bash
python main.py
```

## Como usar

1. Arraste o `.docx` da prova crua para a área pontilhada (ou clique nela
   para escolher pelo explorador de arquivos).
2. Escolha a **Matéria** e o **Bimestre** nos dropdowns.
3. Preencha Professor(a), Série/Turma e Data da prova.
4. O dropdown de **Frase motivacional** já muda sozinho de acordo com a
   matéria escolhida (ver seção abaixo sobre como editar essas frases).
5. Ajuste o nome do arquivo de saída se quiser (por padrão usa o mesmo
   nome do arquivo original).
6. Clique em **FORMATAR PROVA**. O app processa em segundo plano (a
   janela não trava) e mostra o resultado + preview quando terminar.
7. Use **ABRIR PASTA DO ARQUIVO** para ir direto na pasta onde a prova
   foi salva.

## Onde as provas são salvas

Toda prova formatada vai para dentro da pasta `provas/`, organizada
automaticamente assim:

```
provas/
  2026/
    III BIMESTRE/
      PORTUGUÊS/
        PROVA_PORTUGUES_2ANO_A.docx
      MATEMÁTICA/
        PROVA_MATEMATICA_3ANO_B.docx
    IV BIMESTRE/
      ...
```

Se você formatar duas provas com o mesmo nome na mesma pasta, a segunda
NÃO sobrescreve a primeira — o app adiciona `_2`, `_3`, etc.
automaticamente.

## Como editar as frases motivacionais

Abra `frases_motivacionais.json` em qualquer editor de texto. É um
dicionário: cada chave é o nome exato de uma matéria (igual está em
`config.py`), e o valor é uma lista de frases. Exemplo:

```json
{
  "PORTUGUÊS": [
    "SUA DEDICAÇÃO É O QUE TE FAZ VOAR LONGE!",
    "CADA PALAVRA APRENDIDA É UM PASSO A MAIS NA SUA JORNADA!"
  ],
  "MATEMÁTICA": [
    "CADA PROBLEMA RESOLVIDO É UMA VITÓRIA CONQUISTADA!"
  ]
}
```

Só salvar o arquivo já é suficiente — na próxima vez que você trocar de
matéria no dropdown (ou reabrir o app), as frases novas aparecem. Já
vem um punhado de frases de exemplo para as 8 matérias — edite/complete
à vontade.

## Como trocar o cabeçalho decorado (logo + campos)

O app usa `modelo_cabecalho.docx` (na raiz do projeto) como fonte do
cabeçalho decorado — a "pílula" com o logo da escola, Nome do Aluno, Nº,
Série/Turma, Professor(a), Disciplina, Data da Prova, Bimestre e Frase
Motivacional. Esse arquivo **não vem incluído** neste pacote — copie
qualquer prova já diagramada no padrão da escola (com esse cabeçalho) para
`modelo_cabecalho.docx` na raiz do projeto, e o app usa ele
automaticamente a partir da próxima vez que abrir.

Se `modelo_cabecalho.docx` não existir, o app formata a prova normalmente
— só sai sem o cabeçalho decorado (só com a caixa de título laranja).

## Estrutura do projeto

```
main.py                     -> ponto de entrada
config.py                   -> matérias, bimestres, caminhos, cores
formatar_prova.py           -> lógica de formatação do .docx (sem interface)
frases.py                   -> leitura do frases_motivacionais.json
caminhos_saida.py           -> resolve provas/<ano>/<bimestre>/<materia>/arquivo.docx
frases_motivacionais.json   -> frases por matéria (edite à vontade)
modelo_cabecalho.docx       -> (você adiciona) fonte do cabeçalho decorado
ui/
  styles.py                 -> QSS + cores (identidade visual da escola)
  barra_titulo.py           -> BarraTitulo (minimizar + fechar)
  dropzone.py                -> área de arrastar-e-soltar / clique
  preview.py                 -> miniatura da 1ª página (via PyMuPDF)
  worker.py                  -> roda a formatação em thread separada
  janela_principal.py        -> monta a tela inteira
provas/                     -> saída (criada automaticamente)
```

## Limitações conhecidas

- Igual ao script original: quebras de coluna em conteúdo muito extenso
  (poema grande + imagem alta) ainda podem exigir ajuste manual pontual
  no `.docx` gerado.
- O preview mostra só a 1ª página — para conferir a prova inteira, abra
  o `.docx` (ou o `.pdf` que fica salvo do lado, na mesma pasta).
- A detecção automática de disciplina (quando o campo Matéria não bate
  com o título do documento original) segue as mesmas regras do script
  original — sempre confira se o dropdown de Matéria está na opção certa
  antes de formatar.
