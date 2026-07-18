"""
app.py

Interface do usuário, em Streamlit.

Fluxo:
1. Usuário faz upload do .pbix.
2. O arquivo é salvo em um caminho temporário (extracao.py e
   extracao_visuais.py trabalham sobre arquivos em disco).
3. extracao.py lê o modelo semântico (sempre executa — ver instruções do
   projeto), tratando o caso de "thin report" (sem modelo embutido) com uma
   mensagem clara.
4. Três toggles independentes controlam o que entra no documento final:
   - Modelagem (tabelas, medidas, relacionamentos, Power Query)
   - Boas práticas (roda o linter.py só se ativado)
   - Visuais (lê a camada de relatório via extracao_visuais.py só se ativado)
   Qualquer combinação é permitida — inclusive só uma das três, ou nenhuma.
   Um quarto toggle, independente dos três acima, controla uma exportação
   separada: um .docx contendo somente as medidas DAX (nome + código),
   gerado via exportador_medidas_docx.py.
5. gerador.py monta o documento final (Markdown) combinando as seções ativas.
6. O resultado é exibido na tela (pré-visualização) e disponibilizado via
   botão de download. A exportação de medidas em Word, quando ativada, é
   disponibilizada em um botão de download separado.

Execução: `streamlit run src/app.py` (a partir da raiz do projeto).
"""

from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from exportador_medidas import gerar_docx_medidas
from extracao import RelatorioSemModeloError, extrair_modelo
from gerador import renderizar_documentacao

st.set_page_config(
    page_title="Documentador Automático de Modelos Power BI",
    page_icon="",
    layout="wide",
)

st.title("Documentador Automático de Modelos Power BI")
st.caption(
    "Envie um arquivo .pbix para gerar automaticamente a documentação técnica "
    "do modelo semântico e do relatório — 100% local, sem precisar do Power "
    "BI Desktop aberto e sem acesso a nenhum ambiente externo."
)

with st.sidebar:
    st.header("O que incluir no documento")
    incluir_modelagem = st.toggle(
        "Modelagem (tabelas, medidas, relacionamentos)",
        value=True,
        help=(
            "Documentação técnica do modelo semântico: tabelas e colunas, "
            "medidas DAX, relacionamentos e código Power Query (M)."
        ),
    )
    incluir_visuais = st.toggle(
        "Visuais do relatório (páginas, gráficos, KPIs)",
        value=True,
        help=(
            "Lista, por página, os visuais do relatório (gráficos, cartões de "
            "KPI, tabelas, segmentações, etc.) e quais medidas/colunas cada "
            "um usa. Extraído do layout do relatório — formato não "
            "documentado oficialmente pela Microsoft, então a extração é "
            "'best-effort'."
        ),
    )
    incluir_boas_praticas = st.toggle(
        "Boas práticas de modelagem",
        value=True,
        help=(
            "Roda um conjunto de regras sobre o modelo e inclui uma seção "
            "extra com pontos de atenção (medidas sem pasta, relacionamentos "
            "bidirecionais, nomenclatura inconsistente, etc.). Desative para "
            "gerar um documento neutro, sem críticas de modelagem — útil "
            "para entrega direta ao cliente."
        ),
    )
    st.divider()
    st.header("Exportações adicionais")
    incluir_exportacao_medidas_docx = st.toggle(
        "Medidas DAX em Word (.docx)",
        value=False,
        help=(
            "Gera um arquivo .docx separado contendo somente as medidas DAX "
            "do modelo: para cada medida, o nome seguido do código DAX "
            "completo — sem tabelas, colunas, relacionamentos ou qualquer "
            "outra informação. Não substitui a documentação em Markdown "
            "acima; é um arquivo adicional, gerado apenas se este toggle "
            "estiver ativo."
        ),
    )
    st.divider()
    st.caption(
        "Nenhum dado sai da sua máquina: a leitura do .pbix e a geração do "
        "documento acontecem localmente, nesta sessão."
    )
    st.caption(
        "Samuel Villela"
    )

arquivo = st.file_uploader("Arquivo .pbix", type=["pbix"])

if "documento_gerado" not in st.session_state:
    st.session_state.documento_gerado = None
    st.session_state.nome_arquivo_saida = None
    st.session_state.docx_medidas = None
    st.session_state.nome_arquivo_medidas_docx = None

gerar = st.button("Gerar documentação", type="primary", disabled=arquivo is None)

if gerar and arquivo is not None:
    with st.spinner("Lendo o modelo e gerando a documentação..."):
        caminho_temporario = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".pbix", delete=False) as tmp:
                tmp.write(arquivo.getvalue())
                caminho_temporario = Path(tmp.name)

            try:
                modelo = extrair_modelo(caminho_temporario)
            except RelatorioSemModeloError as exc:
                modelo = None
                if incluir_visuais:
                    # Sem modelo semântico (thin report / conexão ao vivo), mas
                    # a camada de visuais é independente e ainda pode ser lida.
                    st.info(
                        f"{exc} Mesmo assim, a documentação dos visuais do "
                        "relatório será gerada, se disponível."
                    )
                else:
                    st.warning(str(exc))
                    raise

            documento = renderizar_documentacao(
                modelo,
                caminho_pbix=caminho_temporario,
                incluir_modelagem=incluir_modelagem,
                incluir_boas_praticas=incluir_boas_praticas,
                incluir_visuais=incluir_visuais,
            )

            st.session_state.documento_gerado = documento
            st.session_state.nome_arquivo_saida = f"{Path(arquivo.name).stem}_documentacao.md"

            # Exportação de medidas em Word — estritamente opt-in (mesmo
            # princípio já aplicado às demais seções) e independente dos três
            # toggles acima: só roda se o toggle específico estiver ativo, e
            # só é possível quando há um modelo semântico embutido (thin
            # reports não têm medidas para exportar).
            if incluir_exportacao_medidas_docx and modelo is not None:
                st.session_state.docx_medidas = gerar_docx_medidas(modelo)
                st.session_state.nome_arquivo_medidas_docx = (
                    f"{Path(arquivo.name).stem}_medidas_dax.docx"
                )
            else:
                st.session_state.docx_medidas = None
                st.session_state.nome_arquivo_medidas_docx = None
                if incluir_exportacao_medidas_docx and modelo is None:
                    st.info(
                        "Não foi possível gerar a exportação de medidas em "
                        "Word: este arquivo não tem um modelo semântico "
                        "embutido (thin report / conexão ao vivo), então não "
                        "há medidas DAX para exportar."
                    )

            if modelo is not None:
                st.success(
                    f"Documentação gerada com sucesso: {modelo.quantidade_tabelas} "
                    f"tabela(s), {modelo.quantidade_medidas} medida(s), "
                    f"{modelo.quantidade_relacionamentos} relacionamento(s)."
                )
            else:
                st.success("Documentação dos visuais gerada com sucesso.")
        except RelatorioSemModeloError:
            st.session_state.documento_gerado = None
            st.session_state.docx_medidas = None
            st.session_state.nome_arquivo_medidas_docx = None
        except Exception as exc:  # noqa: BLE001 — queremos mostrar qualquer erro ao usuário, não só os esperados
            st.session_state.documento_gerado = None
            st.session_state.docx_medidas = None
            st.session_state.nome_arquivo_medidas_docx = None
            st.error(f"Ocorreu um erro inesperado ao processar o arquivo: {exc}")
        finally:
            if caminho_temporario is not None:
                caminho_temporario.unlink(missing_ok=True)

if st.session_state.documento_gerado or st.session_state.docx_medidas:
    colunas_download = st.columns(2)

    with colunas_download[0]:
        if st.session_state.documento_gerado:
            st.download_button(
                "⬇️ Baixar documentação (Markdown)",
                data=st.session_state.documento_gerado,
                file_name=st.session_state.nome_arquivo_saida,
                mime="text/markdown",
            )

    with colunas_download[1]:
        if st.session_state.docx_medidas:
            st.download_button(
                "⬇️ Baixar medidas DAX (Word)",
                data=st.session_state.docx_medidas,
                file_name=st.session_state.nome_arquivo_medidas_docx,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    if st.session_state.documento_gerado:
        st.divider()
        st.subheader("Pré-visualização")
        st.markdown(st.session_state.documento_gerado)