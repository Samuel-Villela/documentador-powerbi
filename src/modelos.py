"""
modelos.py

Estruturas de dados tipadas (pydantic) que representam um modelo semântico
Power BI já extraído e organizado, prontas para serem usadas pelo linter
(linter.py) e pelos templates de geração de documentação (gerador.py).

Manter essas classes desacopladas da extração (extracao.py) é o que permite
trocar a biblioteca de extração no futuro sem impactar o restante do projeto.
"""

from __future__ import annotations

from typing import Dict, List, Optional

from pydantic import BaseModel, Field


class Coluna(BaseModel):
    """Uma coluna de uma tabela do modelo (normal ou calculada)."""

    nome: str
    tipo_dado: str
    e_calculada: bool = False
    expressao: Optional[str] = None
    cardinalidade: Optional[int] = None
    tamanho_dados_bytes: Optional[int] = None


class Medida(BaseModel):
    """Uma medida DAX do modelo."""

    nome: str
    tabela: str
    expressao: str
    pasta_exibicao: Optional[str] = None
    descricao: Optional[str] = None


class Relacionamento(BaseModel):
    """Um relacionamento entre duas tabelas do modelo."""

    tabela_origem: str
    coluna_origem: str
    tabela_destino: str
    coluna_destino: str
    ativo: bool
    cardinalidade: str  # ex.: "M:1", "1:1"
    direcao_filtro: str  # "Única" ou "Ambas"
    integridade_referencial: bool = False


class TabelaCalculada(BaseModel):
    """Uma tabela inteiramente definida por uma expressão DAX (calculated table)."""

    nome: str
    expressao_dax: str


class ConsultaPowerQuery(BaseModel):
    """O código Power Query (M) usado para carregar/transformar uma tabela."""

    tabela: str
    expressao_m: str


class Tabela(BaseModel):
    """Uma tabela "física" do modelo (carregada via Power Query), com suas colunas."""

    nome: str
    colunas: List[Coluna] = Field(default_factory=list)

    @property
    def quantidade_colunas(self) -> int:
        return len(self.colunas)


class CampoVisual(BaseModel):
    """Um campo (coluna ou medida) usado em um visual, em um determinado papel
    (ex.: 'Values', 'Category', 'Y', 'Legend' — varia conforme o tipo de visual)."""

    papel: str
    referencia: str  # referência bruta original (queryRef), útil como fallback
    tabela: Optional[str] = None
    campo: Optional[str] = None
    agregacao_implicita: Optional[str] = None  # ex.: "CountNonNull", quando o próprio visual aplica uma agregação sobre uma coluna
    e_medida: bool = False


class Visual(BaseModel):
    """Um visual individual dentro de uma página do relatório (gráfico, KPI, tabela, etc)."""

    tipo: str
    titulo: Optional[str] = None
    campos: List[CampoVisual] = Field(default_factory=list)


class Pagina(BaseModel):
    """Uma página do relatório, com seus visuais."""

    nome: str
    ordem: int
    visuais: List[Visual] = Field(default_factory=list)

    @property
    def quantidade_visuais(self) -> int:
        return len(self.visuais)


class RelatorioVisual(BaseModel):
    """Agregador da camada de relatório (páginas + visuais) extraída do .pbix."""

    paginas: List[Pagina] = Field(default_factory=list)

    @property
    def quantidade_paginas(self) -> int:
        return len(self.paginas)

    @property
    def quantidade_visuais_total(self) -> int:
        return sum(p.quantidade_visuais for p in self.paginas)


class MetadadosModelo(BaseModel):
    """Metadados gerais do arquivo .pbix e do modelo semântico."""

    nome_arquivo: str
    tamanho_bytes: Optional[int] = None
    versao_power_bi_desktop: Optional[str] = None
    outros: Dict[str, str] = Field(default_factory=dict)


class ModeloPowerBI(BaseModel):
    """Agregador geral: representa o modelo semântico completo extraído de um .pbix."""

    metadados: MetadadosModelo
    tabelas: List[Tabela] = Field(default_factory=list)
    medidas: List[Medida] = Field(default_factory=list)
    relacionamentos: List[Relacionamento] = Field(default_factory=list)
    tabelas_calculadas: List[TabelaCalculada] = Field(default_factory=list)
    power_query: List[ConsultaPowerQuery] = Field(default_factory=list)

    @property
    def quantidade_tabelas(self) -> int:
        return len(self.tabelas)

    @property
    def quantidade_medidas(self) -> int:
        return len(self.medidas)

    @property
    def quantidade_relacionamentos(self) -> int:
        return len(self.relacionamentos)

    def medidas_por_pasta(self) -> Dict[str, List[Medida]]:
        """
        Agrupa as medidas pelo DisplayFolder (pasta de exibição), que é como
        elas aparecem organizadas no painel de campos do Power BI Desktop.

        Medidas sem pasta definida são agrupadas sob a chave "(Sem pasta)".
        O resultado é ordenado alfabeticamente pelo nome da pasta.
        """
        agrupado: Dict[str, List[Medida]] = {}
        for medida in self.medidas:
            chave = medida.pasta_exibicao or "(Sem pasta)"
            agrupado.setdefault(chave, []).append(medida)
        return dict(sorted(agrupado.items(), key=lambda item: item[0]))

    def encontrar_tabela(self, nome: str) -> Optional[Tabela]:
        """Retorna a Tabela com o nome informado, ou None se não existir."""
        for tabela in self.tabelas:
            if tabela.nome == nome:
                return tabela
        return None
