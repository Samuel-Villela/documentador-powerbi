"""
extracao.py

Lê um arquivo .pbix diretamente do disco (via biblioteca `pbixray`, sem
precisar do Power BI Desktop instalado ou aberto) e converte os dados brutos
do modelo semântico nos objetos tipados definidos em modelos.py.

Trata explicitamente o caso de "thin reports" (arquivos .pbix que se conectam
ao vivo a uma fonte externa — Power BI Service ou Analysis Services — e por
isso não têm nenhum modelo embutido no arquivo): nesse caso, uma
RelatorioSemModeloError é levantada com uma mensagem clara, em vez de a
ferramenta falhar silenciosamente ou estourar um erro genérico.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional, Union

import pandas as pd
from pbixray import PBIXRay
from pbixray.exceptions import LiveConnectionError, NoEmbeddedModelError

from modelos import (
    Coluna,
    ConsultaPowerQuery,
    Medida,
    MetadadosModelo,
    ModeloPowerBI,
    Relacionamento,
    Tabela,
    TabelaCalculada,
)


class RelatorioSemModeloError(Exception):
    """
    Levantada quando o .pbix não contém um modelo semântico embutido para
    extrair — por exemplo, "relatórios finos" (thin reports) conectados ao
    vivo a um dataset do Power BI Service ou a um servidor Analysis Services.

    O atributo `tipo_conexao`, quando disponível, indica o tipo de conexão ao
    vivo identificado (ex.: "pbiServiceLive", "analysisServicesDatabaseLive").
    """

    def __init__(self, motivo: str, tipo_conexao: Optional[str] = None):
        self.tipo_conexao = tipo_conexao
        super().__init__(motivo)


def extrair_modelo(caminho_pbix: Union[str, Path]) -> ModeloPowerBI:
    """
    Ponto de entrada principal do módulo: lê o arquivo .pbix em `caminho_pbix`
    e retorna um ModeloPowerBI totalmente populado.

    Levanta RelatorioSemModeloError se o arquivo não tiver modelo embutido.
    """
    caminho_pbix = Path(caminho_pbix)

    try:
        pbix = PBIXRay(str(caminho_pbix))
    except LiveConnectionError as exc:
        raise RelatorioSemModeloError(
            "Este arquivo é um 'relatório fino' (thin report): ele se conecta "
            f"ao vivo a uma fonte externa (tipo de conexão: {exc.connection_type}) "
            "e não tem nenhum modelo semântico embutido no .pbix. Não há nada "
            "para documentar localmente — para gerar a documentação, seria "
            "preciso ter acesso à fonte de dados original (o workspace do "
            "Power BI Service ou o servidor Analysis Services).",
            tipo_conexao=exc.connection_type,
        ) from exc
    except NoEmbeddedModelError as exc:
        raise RelatorioSemModeloError(
            "Não foi encontrado nenhum modelo semântico embutido neste "
            "arquivo .pbix."
        ) from exc

    try:
        return ModeloPowerBI(
            metadados=_extrair_metadados(pbix, caminho_pbix),
            tabelas=_extrair_tabelas(pbix),
            medidas=_extrair_medidas(pbix),
            relacionamentos=_extrair_relacionamentos(pbix),
            tabelas_calculadas=_extrair_tabelas_calculadas(pbix),
            power_query=_extrair_power_query(pbix),
        )
    finally:
        # Libera o arquivo/handles internos usados pelo pbixray.
        pbix.close()


# ---------------------------------------------------------------------------
# Funções internas de extração — cada uma isola o tratamento de uma única
# propriedade do pbixray, incluindo suas particularidades (dataframes vazios,
# colunas ausentes quando não há dados, etc).
# ---------------------------------------------------------------------------


def _extrair_metadados(pbix: PBIXRay, caminho_pbix: Path) -> MetadadosModelo:
    outros: dict[str, str] = {}
    versao_power_bi_desktop: Optional[str] = None

    metadata_df = pbix.metadata
    if metadata_df is not None and not metadata_df.empty:
        for _, linha in metadata_df.iterrows():
            nome = str(linha.get("Name", "")).strip()
            valor = str(linha.get("Value", "")).strip()
            if not nome:
                continue
            if nome == "PBIDesktopVersion":
                versao_power_bi_desktop = valor
            else:
                outros[nome] = valor

    tamanho_bytes: Optional[int] = None
    try:
        tamanho_bytes = int(pbix.size)
    except (TypeError, ValueError):
        tamanho_bytes = None

    return MetadadosModelo(
        nome_arquivo=caminho_pbix.name,
        tamanho_bytes=tamanho_bytes,
        versao_power_bi_desktop=versao_power_bi_desktop,
        outros=outros,
    )


def _extrair_tabelas(pbix: PBIXRay) -> list[Tabela]:
    nomes_tabelas = list(pbix.tables) if pbix.tables is not None else []

    schema_df = pbix.schema
    estatisticas_df = pbix.statistics
    colunas_calculadas_df = pbix.dax_columns

    # Mapa (tabela, coluna) -> (cardinalidade, tamanho_em_bytes), quando disponível.
    estatisticas_por_coluna: dict[tuple[str, str], tuple[Optional[int], Optional[int]]] = {}
    if estatisticas_df is not None and not estatisticas_df.empty:
        for _, linha in estatisticas_df.iterrows():
            chave = (linha["TableName"], linha["ColumnName"])
            estatisticas_por_coluna[chave] = (
                _para_int_ou_none(linha.get("Cardinality")),
                _para_int_ou_none(linha.get("DataSize")),
            )

    # Mapa (tabela, coluna) -> expressão DAX, para colunas calculadas.
    expressao_coluna_calculada: dict[tuple[str, str], str] = {}
    if colunas_calculadas_df is not None and not colunas_calculadas_df.empty:
        for _, linha in colunas_calculadas_df.iterrows():
            chave = (linha["TableName"], linha["ColumnName"])
            expressao_coluna_calculada[chave] = linha.get("Expression") or ""

    tabelas: list[Tabela] = []
    for nome_tabela in nomes_tabelas:
        colunas: list[Coluna] = []

        if schema_df is not None and not schema_df.empty:
            linhas_tabela = schema_df[schema_df["TableName"] == nome_tabela]
            for _, linha in linhas_tabela.iterrows():
                nome_coluna = linha["ColumnName"]
                chave = (nome_tabela, nome_coluna)
                cardinalidade, tamanho_dados = estatisticas_por_coluna.get(chave, (None, None))
                expressao = expressao_coluna_calculada.get(chave)

                colunas.append(
                    Coluna(
                        nome=nome_coluna,
                        tipo_dado=str(linha["PandasDataType"]),
                        e_calculada=chave in expressao_coluna_calculada,
                        expressao=expressao,
                        cardinalidade=cardinalidade,
                        tamanho_dados_bytes=tamanho_dados,
                    )
                )

        tabelas.append(Tabela(nome=nome_tabela, colunas=colunas))

    return tabelas


def _extrair_medidas(pbix: PBIXRay) -> list[Medida]:
    medidas_df = pbix.dax_measures
    if medidas_df is None or medidas_df.empty:
        return []

    medidas: list[Medida] = []
    for _, linha in medidas_df.iterrows():
        medidas.append(
            Medida(
                nome=linha["Name"],
                tabela=linha["TableName"],
                expressao=(linha.get("Expression") or "").strip(),
                pasta_exibicao=_para_str_ou_none(linha.get("DisplayFolder")),
                descricao=_para_str_ou_none(linha.get("Description")),
            )
        )
    return medidas


def _extrair_relacionamentos(pbix: PBIXRay) -> list[Relacionamento]:
    relacionamentos_df = pbix.relationships
    if relacionamentos_df is None or relacionamentos_df.empty:
        return []

    relacionamentos: list[Relacionamento] = []
    for _, linha in relacionamentos_df.iterrows():
        direcao_bruta = linha.get("CrossFilteringBehavior", "Single")
        direcao_filtro = "Ambas" if direcao_bruta == "Both" else "Única"

        relacionamentos.append(
            Relacionamento(
                tabela_origem=linha["FromTableName"],
                coluna_origem=linha["FromColumnName"],
                tabela_destino=linha["ToTableName"],
                coluna_destino=linha["ToColumnName"],
                ativo=bool(linha.get("IsActive", 1)),
                cardinalidade=str(linha.get("Cardinality", "desconhecida")),
                direcao_filtro=direcao_filtro,
                integridade_referencial=bool(linha.get("RelyOnReferentialIntegrity", 0)),
            )
        )
    return relacionamentos


def _extrair_tabelas_calculadas(pbix: PBIXRay) -> list[TabelaCalculada]:
    dax_tables_df = pbix.dax_tables
    if dax_tables_df is None or dax_tables_df.empty:
        return []

    return [
        TabelaCalculada(nome=linha["TableName"], expressao_dax=(linha.get("Expression") or "").strip())
        for _, linha in dax_tables_df.iterrows()
    ]


def _extrair_power_query(pbix: PBIXRay) -> list[ConsultaPowerQuery]:
    power_query_df = pbix.power_query
    if power_query_df is None or power_query_df.empty:
        return []

    return [
        ConsultaPowerQuery(tabela=linha["TableName"], expressao_m=(linha.get("Expression") or "").strip())
        for _, linha in power_query_df.iterrows()
    ]


# ---------------------------------------------------------------------------
# Pequenos utilitários de conversão segura (dataframes podem trazer NaN/None
# em campos opcionais; não queremos propagar "nan" como string para o usuário).
# ---------------------------------------------------------------------------


def _para_int_ou_none(valor) -> Optional[int]:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    try:
        return int(valor)
    except (TypeError, ValueError):
        return None


def _para_str_ou_none(valor) -> Optional[str]:
    if valor is None:
        return None
    if isinstance(valor, float) and pd.isna(valor):
        return None
    texto = str(valor).strip()
    return texto or None
