"""
extracao_visuais.py

Lê a camada de RELATÓRIO do .pbix — páginas e visuais — a partir do arquivo
`Report/Layout`, que fica dentro do contêiner zip do próprio .pbix.

Esta é uma extração DIFERENTE e INDEPENDENTE da extração do modelo semântico
(extracao.py, via `pbixray`): aqui não existe uma biblioteca mantida para nos
ajudar — o `Report/Layout` é um JSON cujo formato foi obtido por engenharia
reversa da comunidade Power BI, não documentado oficialmente pela Microsoft.

Por isso, a extração aqui é deliberadamente "best-effort" e tolerante a
formatos inesperados: um visual ou campo que não seja reconhecido é
representado da forma mais genérica possível (ex.: tipo técnico bruto, sem
título), em vez de fazer a ferramenta inteira falhar.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import List, Optional, Tuple, Union

from modelos import CampoVisual, ModeloPowerBI, Pagina, RelatorioVisual, Visual


class RelatorioSemLayoutError(Exception):
    """
    Levantada quando não é possível extrair páginas/visuais do .pbix — por
    exemplo, porque o arquivo não contém um `Report/Layout` reconhecível, ou
    porque esse layout não pôde ser interpretado como JSON válido.
    """


def extrair_visuais(caminho_pbix: Union[str, Path]) -> RelatorioVisual:
    """
    Ponto de entrada principal do módulo: lê o arquivo .pbix em `caminho_pbix`
    e retorna um RelatorioVisual com as páginas e visuais encontrados.

    Levanta RelatorioSemLayoutError se não houver camada de relatório para
    extrair.
    """
    caminho_pbix = Path(caminho_pbix)
    bruto = _ler_layout_bruto(caminho_pbix)
    dados = _parsear_json(bruto)

    paginas: List[Pagina] = []
    for indice, secao in enumerate(dados.get("sections", [])):
        visuais = [
            visual
            for visual in (
                _extrair_visual(vc) for vc in secao.get("visualContainers", [])
            )
            if visual is not None
        ]
        paginas.append(
            Pagina(
                nome=secao.get("displayName") or f"Página {indice + 1}",
                ordem=indice + 1,
                visuais=visuais,
            )
        )

    return RelatorioVisual(paginas=paginas)


def marcar_medidas_conhecidas(relatorio_visual: RelatorioVisual, modelo: ModeloPowerBI) -> None:
    """
    Cruza os campos usados pelos visuais com as medidas já extraídas do
    modelo (mesma tabela + mesmo nome), marcando `e_medida=True` quando há
    correspondência. Isso permite, na documentação, destacar quando um
    visual usa uma medida DAX (já documentada na seção de modelagem) em vez
    de uma coluna comum.

    Modifica `relatorio_visual` in-place.
    """
    nomes_de_medidas = {(medida.tabela, medida.nome) for medida in modelo.medidas}
    for pagina in relatorio_visual.paginas:
        for visual in pagina.visuais:
            for campo in visual.campos:
                if (campo.tabela, campo.campo) in nomes_de_medidas:
                    campo.e_medida = True


# ---------------------------------------------------------------------------
# Leitura e parsing do arquivo Report/Layout
# ---------------------------------------------------------------------------


def _ler_layout_bruto(caminho_pbix: Path) -> bytes:
    try:
        with zipfile.ZipFile(caminho_pbix) as arquivo_zip:
            nome_layout = next(
                (n for n in arquivo_zip.namelist() if n == "Report/Layout" or n.endswith("/Report/Layout")),
                None,
            )
            if nome_layout is None:
                raise RelatorioSemLayoutError(
                    "Não foi encontrada uma camada de relatório ('Report/Layout') "
                    "neste arquivo .pbix. Pode ser um arquivo somente-modelo, ou "
                    "um formato de relatório não suportado por esta ferramenta."
                )
            return arquivo_zip.read(nome_layout)
    except zipfile.BadZipFile as exc:
        raise RelatorioSemLayoutError(
            "Não foi possível abrir o .pbix como um contêiner válido para "
            "extrair as páginas e visuais do relatório."
        ) from exc


def _parsear_json(bruto: bytes) -> dict:
    texto = _decodificar(bruto)
    try:
        return json.loads(texto)
    except json.JSONDecodeError as exc:
        raise RelatorioSemLayoutError(
            "O layout do relatório está em um formato não reconhecido; não "
            "foi possível interpretar as páginas e visuais."
        ) from exc


def _decodificar(bruto: bytes) -> str:
    """O Report/Layout normalmente vem em UTF-16 LE (com ou sem BOM), mas
    tentamos algumas codificações alternativas antes de desistir."""
    for codificacao in ("utf-16", "utf-8-sig", "utf-8"):
        try:
            return bruto.decode(codificacao)
        except UnicodeDecodeError:
            continue
    return bruto.decode("utf-8", errors="replace")


# ---------------------------------------------------------------------------
# Interpretação de cada visual individual
# ---------------------------------------------------------------------------

_MAPA_TIPOS_VISUAL = {
    "card": "Cartão (KPI de valor único)",
    "multiRowCard": "Cartão multi-linha",
    "kpi": "KPI (indicador com meta)",
    "gauge": "Medidor (gauge)",
    "clusteredColumnChart": "Gráfico de colunas agrupadas",
    "clusteredBarChart": "Gráfico de barras agrupadas",
    "columnChart": "Gráfico de colunas empilhadas",
    "barChart": "Gráfico de barras empilhadas",
    "lineChart": "Gráfico de linhas",
    "lineStackedColumnComboChart": "Gráfico combinado (linhas e colunas)",
    "lineClusteredColumnComboChart": "Gráfico combinado (linhas e colunas agrupadas)",
    "pieChart": "Gráfico de pizza",
    "donutChart": "Gráfico de rosca",
    "treemap": "Mapa de árvore (treemap)",
    "scatterChart": "Gráfico de dispersão",
    "areaChart": "Gráfico de área",
    "waterfallChart": "Gráfico de cascata",
    "funnel": "Gráfico de funil",
    "ribbonChart": "Gráfico de faixas (ribbon)",
    "tableEx": "Tabela",
    "pivotTable": "Matriz (tabela dinâmica)",
    "slicer": "Segmentação de dados (filtro)",
    "map": "Mapa",
    "filledMap": "Mapa preenchido (coroplético)",
    "shapeMap": "Mapa de formas",
    "textbox": "Caixa de texto",
    "image": "Imagem",
    "shape": "Forma",
    "actionButton": "Botão de ação",
}


def _nome_amigavel_tipo(tipo_bruto: Optional[str]) -> str:
    if not tipo_bruto:
        return "Tipo de visual não identificado"
    return _MAPA_TIPOS_VISUAL.get(tipo_bruto, f"Outro ({tipo_bruto})")


def _extrair_visual(container: dict) -> Optional[Visual]:
    config_bruto = container.get("config")
    if not config_bruto:
        return None

    try:
        config = json.loads(config_bruto)
    except (json.JSONDecodeError, TypeError):
        return None

    single = config.get("singleVisual")
    if single is None:
        # Pode ser um grupo de visuais (visualGroup) ou outro contêiner
        # estrutural sem um tipo de visual próprio — fora do escopo desta v1.
        return None

    return Visual(
        tipo=_nome_amigavel_tipo(single.get("visualType")),
        titulo=_extrair_titulo(single),
        campos=_extrair_campos(single),
    )


def _extrair_titulo(single: dict) -> Optional[str]:
    try:
        valor = single["vcObjects"]["title"][0]["properties"]["text"]["expr"]["Literal"]["Value"]
    except (KeyError, IndexError, TypeError):
        return None
    if not isinstance(valor, str):
        return None
    return valor.strip("'\"") or None


_PADRAO_AGREGACAO_IMPLICITA = re.compile(r"^([A-Za-z][A-Za-z0-9_]*)\((.+)\)$")


def _extrair_campos(single: dict) -> List[CampoVisual]:
    campos: List[CampoVisual] = []
    for papel, itens in (single.get("projections") or {}).items():
        for item in itens:
            referencia = item.get("queryRef") or ""
            tabela, campo, agregacao = _interpretar_referencia(referencia)
            campos.append(
                CampoVisual(
                    papel=papel,
                    referencia=referencia,
                    tabela=tabela,
                    campo=campo,
                    agregacao_implicita=agregacao,
                )
            )
    return campos


def _interpretar_referencia(referencia: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Tenta separar uma referência de campo (queryRef) em tabela/campo, e
    identificar se há uma agregação aplicada diretamente pelo visual (ex.:
    "CountNonNull(Product.Color)" -> agregação "CountNonNull" sobre
    "Product.Color").

    Referências que não seguem nenhum desses padrões são mantidas apenas em
    `referencia` (já presente em CampoVisual), sem tabela/campo separados.
    """
    texto = referencia
    agregacao = None

    correspondencia = _PADRAO_AGREGACAO_IMPLICITA.match(texto)
    if correspondencia:
        agregacao = correspondencia.group(1)
        texto = correspondencia.group(2)

    if "." in texto and "(" not in texto:
        tabela, _, campo = texto.rpartition(".")
        return (tabela or None), (campo or None), agregacao

    return None, None, agregacao
