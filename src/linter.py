"""
linter.py

Regras de boas práticas de modelagem, executadas OPCIONALMENTE sobre um
ModeloPowerBI já estruturado (só quando o usuário ativa o toggle na
interface — ver app.py).

Cada regra é uma função independente que recebe o ModeloPowerBI e devolve uma
lista de PontoAtencao. Isso deixa fácil adicionar, remover ou desativar regras
individualmente no futuro (ver seção "Possíveis evoluções futuras" das
instruções do projeto).

Cada PontoAtencao inclui, além da mensagem, uma breve explicação do "porquê"
aquilo é considerado um ponto de atenção — para que o relatório também sirva
como material de aprendizado para analistas menos experientes.
"""

from __future__ import annotations

import re
from enum import Enum
from typing import Callable, List, Optional

from pydantic import BaseModel

from modelos import Coluna, Medida, ModeloPowerBI, Relacionamento, Tabela


class Severidade(str, Enum):
    BAIXA = "baixa"
    MEDIA = "média"
    ALTA = "alta"


class PontoAtencao(BaseModel):
    regra: str
    severidade: Severidade
    mensagem: str
    localizacao: Optional[str] = None


# ---------------------------------------------------------------------------
# Regra 1 — Medidas sem DisplayFolder
# ---------------------------------------------------------------------------


def _regra_medidas_sem_pasta(modelo: ModeloPowerBI) -> List[PontoAtencao]:
    """
    Medidas sem `DisplayFolder` definido ficam soltas na raiz do painel de
    campos, dificultando a organização do modelo à medida que o número de
    medidas cresce.
    """
    pontos = []
    for medida in modelo.medidas:
        if not medida.pasta_exibicao:
            pontos.append(
                PontoAtencao(
                    regra="Medidas sem pasta de exibição (DisplayFolder)",
                    severidade=Severidade.BAIXA,
                    mensagem=(
                        f"A medida '{medida.nome}' não tem uma pasta de exibição "
                        "definida. Isso dificulta a organização do painel de "
                        "campos conforme o modelo cresce."
                    ),
                    localizacao=f"Tabela: {medida.tabela}, Medida: {medida.nome}",
                )
            )
    return pontos


# ---------------------------------------------------------------------------
# Regra 2 — Nome sugere formatação, mas a expressão não usa FORMAT()
# ---------------------------------------------------------------------------

_SUFIXOS_QUE_SUGEREM_FORMATACAO = ("%", "r$", "us$", "€", "pct", "percentual")


def _regra_nome_sugere_formatacao(modelo: ModeloPowerBI) -> List[PontoAtencao]:
    """
    Quando o nome de uma medida sugere um formato específico (termina em "%",
    menciona "R$", etc.), mas a expressão não usa FORMAT(), o valor exibido
    depende inteiramente da formatação configurada manualmente na medida — o
    que é fácil de perder ao duplicar/mover a medida, e não fica evidente só
    de olhar o código DAX.
    """
    pontos = []
    for medida in modelo.medidas:
        nome_normalizado = medida.nome.strip().lower()
        sugere_formatacao = any(
            nome_normalizado.endswith(sufixo) or sufixo in nome_normalizado
            for sufixo in _SUFIXOS_QUE_SUGEREM_FORMATACAO
        )
        usa_format = "format(" in medida.expressao.lower()

        if sugere_formatacao and not usa_format:
            pontos.append(
                PontoAtencao(
                    regra="Nome sugere formatação, mas expressão não usa FORMAT()",
                    severidade=Severidade.BAIXA,
                    mensagem=(
                        f"O nome da medida '{medida.nome}' sugere um formato "
                        "específico (percentual, moeda, etc.), mas a expressão "
                        "não usa FORMAT(). Verifique se a formatação está "
                        "garantida via propriedade da medida, e considere "
                        "deixar isso explícito na expressão."
                    ),
                    localizacao=f"Tabela: {medida.tabela}, Medida: {medida.nome}",
                )
            )
    return pontos


# ---------------------------------------------------------------------------
# Regra 3 — Relacionamentos bidirecionais
# ---------------------------------------------------------------------------


def _regra_relacionamentos_bidirecionais(modelo: ModeloPowerBI) -> List[PontoAtencao]:
    """
    Relacionamentos com filtro cruzado em "ambas as direções" são uma fonte
    comum de ambiguidade e lentidão em modelos Power BI: eles permitem que o
    filtro se propague nos dois sentidos entre as tabelas, o que pode gerar
    contextos de filtro inesperados e, em modelos maiores, prejudicar a
    performance. Muitas vezes o mesmo resultado pode ser obtido com um
    relacionamento de direção única combinado a uma medida com CALCULATE/
    CROSSFILTER pontual, de forma mais previsível.
    """
    pontos = []
    for rel in modelo.relacionamentos:
        if rel.direcao_filtro == "Ambas":
            pontos.append(
                PontoAtencao(
                    regra="Relacionamentos bidirecionais",
                    severidade=Severidade.MEDIA,
                    mensagem=(
                        f"O relacionamento entre '{rel.tabela_origem}' e "
                        f"'{rel.tabela_destino}' está configurado com filtro "
                        "cruzado em ambas as direções. Avalie se ele realmente "
                        "precisa ser bidirecional, ou se pode ser unidirecional "
                        "(com CROSSFILTER aplicado pontualmente onde for "
                        "necessário)."
                    ),
                    localizacao=(
                        f"{rel.tabela_origem}.{rel.coluna_origem} -> "
                        f"{rel.tabela_destino}.{rel.coluna_destino}"
                    ),
                )
            )
    return pontos


# ---------------------------------------------------------------------------
# Regra 4 — Colunas calculadas que poderiam ser medidas
# ---------------------------------------------------------------------------

_FUNCOES_DE_AGREGACAO = (
    "sum(", "sumx(", "count(", "counta(", "countx(", "countrows(",
    "average(", "averagex(", "min(", "minx(", "max(", "maxx(",
    "calculate(", "distinctcount(",
)


def _regra_colunas_calculadas_candidatas_a_medida(modelo: ModeloPowerBI) -> List[PontoAtencao]:
    """
    Colunas calculadas são materializadas em disco: o valor é calculado uma
    vez por linha e armazenado, aumentando o tamanho do modelo. Quando a
    expressão de uma coluna calculada usa funções de agregação (SUM, COUNT,
    CALCULATE, etc.), muitas vezes o mesmo resultado poderia ser obtido de
    forma mais eficiente com uma medida, calculada sob demanda em tempo de
    consulta em vez de armazenada por linha.
    """
    pontos = []
    for tabela in modelo.tabelas:
        for coluna in tabela.colunas:
            if not coluna.e_calculada or not coluna.expressao:
                continue
            expressao_normalizada = coluna.expressao.lower()
            if any(funcao in expressao_normalizada for funcao in _FUNCOES_DE_AGREGACAO):
                pontos.append(
                    PontoAtencao(
                        regra="Colunas calculadas candidatas a virar medida",
                        severidade=Severidade.MEDIA,
                        mensagem=(
                            f"A coluna calculada '{coluna.nome}' (tabela "
                            f"'{tabela.nome}') usa função(ões) de agregação em "
                            "sua expressão. Colunas calculadas são armazenadas "
                            "linha a linha; se o valor não depende do contexto "
                            "de cada linha individual, considere transformá-la "
                            "em uma medida para reduzir o tamanho do modelo."
                        ),
                        localizacao=f"Tabela: {tabela.nome}, Coluna: {coluna.nome}",
                    )
                )
    return pontos


# ---------------------------------------------------------------------------
# Regra 5 — Tabelas órfãs (sem nenhum relacionamento)
# ---------------------------------------------------------------------------


def _regra_tabelas_orfas(modelo: ModeloPowerBI) -> List[PontoAtencao]:
    """
    Uma tabela sem nenhum relacionamento com o restante do modelo não
    participa do filtro cruzado entre visuais, o que normalmente indica uma
    tabela esquecida, um relacionamento que falhou ao ser criado, ou uma
    tabela usada apenas para cálculos isolados (o que deveria estar
    documentado explicitamente).
    """
    tabelas_com_relacionamento = set()
    for rel in modelo.relacionamentos:
        tabelas_com_relacionamento.add(rel.tabela_origem)
        tabelas_com_relacionamento.add(rel.tabela_destino)

    pontos = []
    for tabela in modelo.tabelas:
        if tabela.nome not in tabelas_com_relacionamento:
            pontos.append(
                PontoAtencao(
                    regra="Tabelas órfãs (sem relacionamento)",
                    severidade=Severidade.ALTA,
                    mensagem=(
                        f"A tabela '{tabela.nome}' não tem nenhum "
                        "relacionamento com as demais tabelas do modelo. "
                        "Verifique se isso é intencional (ex.: tabela de "
                        "parâmetros/cálculo isolado) ou se um relacionamento "
                        "deveria ter sido criado."
                    ),
                    localizacao=f"Tabela: {tabela.nome}",
                )
            )
    return pontos


# ---------------------------------------------------------------------------
# Regra 6 — Inconsistência de nomenclatura
# ---------------------------------------------------------------------------

_PADRAO_SNAKE_CASE = re.compile(r"^[a-z0-9]+(_[a-z0-9]+)*$")
_PADRAO_PASCAL_CASE = re.compile(r"^([A-Z][a-z0-9]*)+$")
_PADRAO_TITULO_COM_ESPACOS = re.compile(r"^[A-Za-z0-9]+( [A-Za-z0-9]+)+$")


def _classificar_convencao(nome: str) -> str:
    """Classifica um identificador em uma convenção de nomenclatura conhecida."""
    if _PADRAO_SNAKE_CASE.match(nome):
        return "snake_case"
    if _PADRAO_PASCAL_CASE.match(nome):
        return "PascalCase"
    if _PADRAO_TITULO_COM_ESPACOS.match(nome):
        return "Título Com Espaços"
    return "outro"


def _regra_inconsistencia_nomenclatura(modelo: ModeloPowerBI) -> List[PontoAtencao]:
    """
    Misturar convenções de nomenclatura (ex.: `PascalCase` para algumas
    tabelas e `snake_case` para outras) deixa o modelo com aparência
    inconsistente e dificulta a previsibilidade de nomes ao escrever DAX —
    o analista precisa lembrar, caso a caso, qual convenção foi usada onde.
    """
    pontos: List[PontoAtencao] = []

    def _verificar(nomes: List[str], rotulo: str):
        convencoes = {nome: _classificar_convencao(nome) for nome in nomes if nome}
        convencoes_encontradas = {c for c in convencoes.values() if c != "outro"}
        if len(convencoes_encontradas) > 1:
            exemplos = ", ".join(f"'{nome}' ({convencao})" for nome, convencao in list(convencoes.items())[:5])
            pontos.append(
                PontoAtencao(
                    regra="Inconsistência de nomenclatura",
                    severidade=Severidade.BAIXA,
                    mensagem=(
                        f"Os nomes de {rotulo} misturam mais de uma convenção "
                        f"de nomenclatura ({', '.join(sorted(convencoes_encontradas))}). "
                        f"Exemplos: {exemplos}."
                    ),
                    localizacao=None,
                )
            )

    _verificar([t.nome for t in modelo.tabelas], "tabelas")
    _verificar([m.nome for m in modelo.medidas], "medidas")

    return pontos


# ---------------------------------------------------------------------------
# Regra 7 — Medidas muito longas ou muito aninhadas
# ---------------------------------------------------------------------------

_LIMITE_CARACTERES = 400
_LIMITE_PROFUNDIDADE_PARENTESES = 6


def _profundidade_maxima_parenteses(expressao: str) -> int:
    profundidade = 0
    profundidade_maxima = 0
    for caractere in expressao:
        if caractere == "(":
            profundidade += 1
            profundidade_maxima = max(profundidade_maxima, profundidade)
        elif caractere == ")":
            profundidade = max(0, profundidade - 1)
    return profundidade_maxima


def _regra_medidas_longas_ou_aninhadas(modelo: ModeloPowerBI) -> List[PontoAtencao]:
    """
    Medidas muito longas ou com muitos níveis de aninhamento de funções são
    difíceis de ler, revisar e depurar. Quebrá-las em medidas auxiliares
    (medidas intermediárias que a medida final referencia) costuma facilitar
    a manutenção e permite reaproveitar os cálculos intermediários em outras
    medidas.
    """
    pontos = []
    for medida in modelo.medidas:
        tamanho = len(medida.expressao)
        profundidade = _profundidade_maxima_parenteses(medida.expressao)

        if tamanho > _LIMITE_CARACTERES or profundidade > _LIMITE_PROFUNDIDADE_PARENTESES:
            motivo = []
            if tamanho > _LIMITE_CARACTERES:
                motivo.append(f"{tamanho} caracteres")
            if profundidade > _LIMITE_PROFUNDIDADE_PARENTESES:
                motivo.append(f"{profundidade} níveis de aninhamento de parênteses")

            pontos.append(
                PontoAtencao(
                    regra="Medidas muito longas ou aninhadas",
                    severidade=Severidade.BAIXA,
                    mensagem=(
                        f"A medida '{medida.nome}' é extensa ({' e '.join(motivo)}). "
                        "Considere quebrá-la em medidas auxiliares para facilitar "
                        "a leitura e a manutenção."
                    ),
                    localizacao=f"Tabela: {medida.tabela}, Medida: {medida.nome}",
                )
            )
    return pontos


# ---------------------------------------------------------------------------
# Execução de todas as regras
# ---------------------------------------------------------------------------

_TODAS_AS_REGRAS: List[Callable[[ModeloPowerBI], List[PontoAtencao]]] = [
    _regra_medidas_sem_pasta,
    _regra_nome_sugere_formatacao,
    _regra_relacionamentos_bidirecionais,
    _regra_colunas_calculadas_candidatas_a_medida,
    _regra_tabelas_orfas,
    _regra_inconsistencia_nomenclatura,
    _regra_medidas_longas_ou_aninhadas,
]

_ORDEM_SEVERIDADE = {Severidade.ALTA: 0, Severidade.MEDIA: 1, Severidade.BAIXA: 2}


def analisar_boas_praticas(modelo: ModeloPowerBI) -> List[PontoAtencao]:
    """
    Executa todas as regras de boas práticas sobre o modelo e retorna a lista
    combinada de pontos de atenção, ordenada por severidade (alta -> baixa) e,
    dentro de cada severidade, pelo nome da regra.
    """
    pontos: List[PontoAtencao] = []
    for regra in _TODAS_AS_REGRAS:
        pontos.extend(regra(modelo))

    pontos.sort(key=lambda p: (_ORDEM_SEVERIDADE[p.severidade], p.regra))
    return pontos
