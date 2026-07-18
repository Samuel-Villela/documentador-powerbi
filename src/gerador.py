"""
gerador.py

Renderiza o documento final em Markdown a partir de um ModeloPowerBI,
combinando três seções INDEPENDENTES, cada uma controlada pelo seu próprio
toggle (o usuário pode ativar qualquer combinação delas):

- `incluir_modelagem`     -> tabelas, colunas, medidas, relacionamentos, Power Query.
- `incluir_boas_praticas` -> roda o linter (linter.py) e mostra os pontos de atenção.
- `incluir_visuais`       -> lê a camada de relatório (extracao_visuais.py) e
                             mostra páginas, visuais e os campos/medidas que
                             cada um usa.

Regra de design (mantida das instruções originais do projeto, agora
generalizada para 3 toggles em vez de 1): cada seção só é processada/rodada
se o seu toggle estiver ativo, e cada uma vive em seu próprio template
parcial — nunca como condicionais espalhadas dentro de um template único.
O template mestre (`documento.md.j2`) apenas decide, em um único lugar, quais
seções incluir; toda a lógica de cada seção fica isolada em seu próprio
arquivo `.j2`.
"""

from __future__ import annotations

from datetime import datetime
from itertools import groupby
from pathlib import Path
from typing import Dict, List, Optional, Union

from jinja2 import Environment, FileSystemLoader, select_autoescape

from extracao_visuais import RelatorioSemLayoutError, extrair_visuais, marcar_medidas_conhecidas
from linter import PontoAtencao, analisar_boas_praticas
from modelos import ModeloPowerBI, RelatorioVisual

TEMPLATES_DIR = Path(__file__).resolve().parent.parent / "templates"
TEMPLATE_MESTRE = "documento.md.j2"


def _criar_ambiente_jinja() -> Environment:
    return Environment(
        loader=FileSystemLoader(str(TEMPLATES_DIR)),
        autoescape=select_autoescape(enabled_extensions=(), default=False),
        trim_blocks=True,
        lstrip_blocks=True,
    )


def _agrupar_pontos_por_regra(pontos: List[PontoAtencao]) -> Dict[str, List[PontoAtencao]]:
    """Agrupa os pontos de atenção pelo nome da regra, preservando a ordem
    (que já vem por severidade, definida em linter.analisar_boas_praticas)."""
    agrupado: Dict[str, List[PontoAtencao]] = {}
    pontos_ordenados = sorted(pontos, key=lambda p: p.regra)
    for regra, grupo in groupby(pontos_ordenados, key=lambda p: p.regra):
        agrupado[regra] = list(grupo)
    return agrupado


def renderizar_documentacao(
    modelo: Optional[ModeloPowerBI],
    caminho_pbix: Union[str, Path, None] = None,
    incluir_modelagem: bool = True,
    incluir_boas_praticas: bool = True,
    incluir_visuais: bool = True,
    gerado_em: Optional[datetime] = None,
) -> str:
    """
    Gera o documento final (Markdown), combinando as três seções conforme os
    toggles ativados.

    `modelo` pode ser None quando não há modelo semântico embutido no .pbix
    (thin reports / conexão ao vivo) — nesse caso, `incluir_modelagem` e
    `incluir_boas_praticas` são automaticamente desativados (não há o que
    documentar ou analisar), e apenas a seção de visuais pode ser gerada, se
    solicitada e se `caminho_pbix` for informado.

    `caminho_pbix` só é necessário quando `incluir_visuais=True` (para ler o
    `Report/Layout` do arquivo original); nos demais casos pode ser omitido.
    """
    gerado_em = gerado_em or datetime.now()

    notas: List[str] = []
    if modelo is None:
        if incluir_modelagem or incluir_boas_praticas:
            notas.append(
                "Este arquivo não tem um modelo semântico embutido (thin "
                "report / conexão ao vivo), então as seções de modelagem e "
                "de boas práticas não puderam ser geradas — apenas a "
                "camada de visuais, quando disponível."
            )
        incluir_modelagem = False
        incluir_boas_praticas = False

    nome_arquivo = (
        modelo.metadados.nome_arquivo
        if modelo is not None
        else (Path(caminho_pbix).name if caminho_pbix is not None else "arquivo desconhecido")
    )

    pontos_atencao: List[PontoAtencao] = []
    pontos_por_regra: Dict[str, List[PontoAtencao]] = {}
    if incluir_boas_praticas and modelo is not None:
        pontos_atencao = analisar_boas_praticas(modelo)
        pontos_por_regra = _agrupar_pontos_por_regra(pontos_atencao)

    relatorio_visual: Optional[RelatorioVisual] = None
    erro_visuais: Optional[str] = None
    if incluir_visuais:
        if caminho_pbix is None:
            raise ValueError(
                "caminho_pbix é obrigatório quando incluir_visuais=True, para "
                "ler a camada de relatório (Report/Layout) do arquivo original."
            )
        try:
            relatorio_visual = extrair_visuais(caminho_pbix)
            if modelo is not None:
                marcar_medidas_conhecidas(relatorio_visual, modelo)
        except RelatorioSemLayoutError as exc:
            # Não é um erro fatal: o restante do documento (modelagem/boas
            # práticas) continua sendo gerado normalmente; só avisamos que a
            # seção de visuais não pôde ser preenchida, e por quê.
            erro_visuais = str(exc)

    secoes_incluidas = []
    if incluir_modelagem:
        secoes_incluidas.append("Modelagem do modelo")
    if incluir_visuais:
        secoes_incluidas.append("Visuais do relatório")
    if incluir_boas_praticas:
        secoes_incluidas.append("Boas práticas")

    ambiente = _criar_ambiente_jinja()
    template = ambiente.get_template(TEMPLATE_MESTRE)

    return template.render(
        modelo=modelo,
        nome_arquivo=nome_arquivo,
        gerado_em=gerado_em,
        incluir_modelagem=incluir_modelagem,
        incluir_boas_praticas=incluir_boas_praticas,
        incluir_visuais=incluir_visuais,
        pontos_atencao=pontos_atencao,
        pontos_por_regra=pontos_por_regra,
        relatorio_visual=relatorio_visual,
        erro_visuais=erro_visuais,
        secoes_incluidas=secoes_incluidas,
        notas=notas,
    )
