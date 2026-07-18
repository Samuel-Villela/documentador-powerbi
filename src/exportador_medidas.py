"""
exportador_medidas_docx.py

Gera um documento .docx contendo SOMENTE as medidas DAX do modelo: para cada
medida, o nome seguido do código DAX completo — sem tabelas, colunas,
relacionamentos, Power Query ou qualquer outra seção da documentação.

Esta é uma saída independente da documentação em Markdown (gerador.py /
templates/*.md.j2): não gera um Markdown e converte, e não reaproveita o
Jinja2. Em vez disso, usa a mesma fonte de dados já extraída pelo restante da
ferramenta — `ModeloPowerBI.medidas` (ver extracao.py / modelos.py) — apenas
trocando a camada de saída para .docx. Assim, qualquer melhoria futura na
extração de medidas (extracao.py) já beneficia esta exportação automaticamente,
sem duplicar lógica.

Deliberadamente sem formatação ou destaque de sintaxe: cada medida é
renderizada como nome (em negrito) + código DAX em fonte monoespaçada simples,
conforme solicitado.
"""

from __future__ import annotations

import io

from docx import Document
from docx.shared import Pt

from modelos import ModeloPowerBI

_FONTE_CODIGO = "Consolas"


def gerar_docx_medidas(modelo: ModeloPowerBI) -> bytes:
    """
    Gera, em memória, um documento .docx contendo apenas as medidas DAX de
    `modelo` e retorna seu conteúdo como bytes (pronto para uso em
    `st.download_button` ou para ser salvo em disco).

    Para cada medida: um parágrafo com o nome (negrito) seguido de um
    parágrafo com a expressão DAX completa (fonte monoespaçada, sem
    destaque de sintaxe). Se o modelo não tiver nenhuma medida, o documento
    gerado contém apenas uma mensagem indicando isso.
    """
    documento = Document()

    if not modelo.medidas:
        documento.add_paragraph("Nenhuma medida DAX encontrada neste modelo.")
    else:
        ultima = len(modelo.medidas) - 1
        for indice, medida in enumerate(modelo.medidas):
            _adicionar_medida(documento, medida.nome, medida.expressao)
            if indice != ultima:
                documento.add_paragraph()  # linha em branco separando as medidas

    buffer = io.BytesIO()
    documento.save(buffer)
    return buffer.getvalue()


def _adicionar_medida(documento: Document, nome: str, expressao_dax: str) -> None:
    paragrafo_nome = documento.add_paragraph()
    execucao_nome = paragrafo_nome.add_run(nome)
    execucao_nome.bold = True
    execucao_nome.font.size = Pt(12)

    # O código DAX pode ter múltiplas linhas; cada quebra de linha vira um
    # novo "run" quebrado (add_break), já que o docx não interpreta "\n".
    paragrafo_dax = documento.add_paragraph()
    linhas = (expressao_dax or "").split("\n")
    for indice_linha, linha in enumerate(linhas):
        execucao_dax = paragrafo_dax.add_run(linha)
        execucao_dax.font.name = _FONTE_CODIGO
        execucao_dax.font.size = Pt(10)
        if indice_linha != len(linhas) - 1:
            execucao_dax.add_break()