"""
extractor.py
Lê o histograma mensal da Ademicon e extrai os dados de cada grupo.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta


TAXA_ADM_PADRAO = 0.24  # 24% para grupos imóveis padrão


def extrair_base_grupos(hist_path: str, meses_historico: int = 6) -> pd.DataFrame:
    """
    Lê o histograma e retorna um DataFrame com uma linha por grupo ativo,
    com todas as métricas necessárias para o motor de cálculo.

    Retorna colunas:
      grupo, participantes, prazo_original, taxa_adm, credito,
      prazo_atual, parcela_lancamento, parcela_atual,
      lance_medio_parcelas, lance_R, lance_embutido_R, lance_livre_R,
      ss_novo, nova_parcela, contemp_por_mes_ll, tipos_lance
    """
    xl = pd.ExcelFile(hist_path)

    # --- 1. Histograma principal ---
    df_hist = pd.read_excel(hist_path, sheet_name="Histograma", header=None)
    df_hist.columns = df_hist.iloc[7]
    df_hist = df_hist.iloc[8:].reset_index(drop=True)

    df_hist["ASSEMBLEIA"] = pd.to_datetime(df_hist["ASSEMBLEIA"], errors="coerce")
    df_hist = df_hist.dropna(subset=["ASSEMBLEIA", "GRUPO"])
    df_hist["GRUPO"] = pd.to_numeric(df_hist["GRUPO"], errors="coerce")
    df_hist = df_hist.dropna(subset=["GRUPO"])
    df_hist["GRUPO"] = df_hist["GRUPO"].astype(int)

    # Filtrar apenas os últimos N meses
    data_corte = df_hist["ASSEMBLEIA"].max() - pd.DateOffset(months=meses_historico)
    df_recente = df_hist[df_hist["ASSEMBLEIA"] >= data_corte].copy()

    # Apenas confirmados
    df_conf = df_recente[df_recente["STATUS"] == "CONFIRMADO"].copy()
    df_conf["Nª PARTICIPANTES"] = pd.to_numeric(df_conf["Nª PARTICIPANTES"], errors="coerce")
    df_conf["PRAZO DO GRUPO"] = pd.to_numeric(df_conf["PRAZO DO GRUPO"], errors="coerce")
    df_conf["QUANTIDADE DE PARCELAS OFERTADAS"] = pd.to_numeric(
        df_conf["QUANTIDADE DE PARCELAS OFERTADAS"], errors="coerce"
    )

    # Separar tipos de lance
    tipos_lance = ["Lance Livre", "Lance Fixo", "Lance Limitado", "Lance Fidelidade"]
    df_lance = df_conf[df_conf["TIPO DE CONTEMPLAÇÃO"].isin(tipos_lance)].copy()
    df_sorteio = df_conf[df_conf["TIPO DE CONTEMPLAÇÃO"].str.contains("Sorteio", na=False)].copy()

    # Média de parcelas ofertadas por grupo e tipo
    lance_stats = (
        df_lance.groupby(["GRUPO", "TIPO DE CONTEMPLAÇÃO"])["QUANTIDADE DE PARCELAS OFERTADAS"]
        .agg(["mean", "count"])
        .reset_index()
        .rename(columns={"mean": "media_parcelas", "count": "qtd"})
    )

    # Contemplações mensais por grupo (lance livre + sorteio)
    df_hist_mes = df_recente.copy()
    df_hist_mes["MES"] = df_hist_mes["ASSEMBLEIA"].dt.to_period("M")
    contemp_mes = (
        df_hist_mes[df_hist_mes["STATUS"] == "CONFIRMADO"]
        .groupby(["GRUPO", "MES"])
        .size()
        .reset_index(name="contemp")
        .groupby("GRUPO")["contemp"]
        .mean()
        .reset_index(name="contemp_por_mes")
    )

    contemp_ll_mes = (
        df_hist_mes[
            (df_hist_mes["STATUS"] == "CONFIRMADO")
            & (df_hist_mes["TIPO DE CONTEMPLAÇÃO"] == "Lance Livre")
        ]
        .groupby(["GRUPO", "MES"])
        .size()
        .reset_index(name="contemp")
        .groupby("GRUPO")["contemp"]
        .mean()
        .reset_index(name="contemp_ll_mes")
    )

    # Info base por grupo (participantes, prazo)
    info_grupo = (
        df_conf.groupby("GRUPO")
        .agg(
            participantes=("Nª PARTICIPANTES", "first"),
            prazo_original=("PRAZO DO GRUPO", "first"),
        )
        .reset_index()
    )

    # --- 2. Regras: extrair crédito aproximado e prazo ---
    df_regras = pd.read_excel(hist_path, sheet_name="Regras", header=None)
    grupo_prazo = _extrair_prazo_das_regras(df_regras)

    # --- 3. Montar tabela de lances por grupo ---
    # Pivot: lance livre médio e embutido médio
    lance_livre = lance_stats[lance_stats["TIPO DE CONTEMPLAÇÃO"] == "Lance Livre"][
        ["GRUPO", "media_parcelas"]
    ].rename(columns={"media_parcelas": "lance_livre_parcelas"})

    lance_embutido = lance_stats[
        lance_stats["TIPO DE CONTEMPLAÇÃO"].isin(["Lance Fixo", "Lance Limitado"])
    ].groupby("GRUPO")["media_parcelas"].mean().reset_index().rename(
        columns={"media_parcelas": "lance_embutido_parcelas"}
    )

    lance_fidelidade = lance_stats[
        lance_stats["TIPO DE CONTEMPLAÇÃO"] == "Lance Fidelidade"
    ][["GRUPO", "media_parcelas"]].rename(
        columns={"media_parcelas": "lance_fidelidade_parcelas"}
    )

    # Tipos de lance disponíveis por grupo
    tipos_por_grupo = (
        lance_stats.groupby("GRUPO")["TIPO DE CONTEMPLAÇÃO"]
        .apply(lambda x: ", ".join(sorted(x.unique())))
        .reset_index(name="tipos_lance")
    )

    # --- 4. Juntar tudo ---
    base = info_grupo.copy()
    base = base.merge(grupo_prazo, on="GRUPO", how="left")
    base = base.merge(lance_livre, on="GRUPO", how="left")
    base = base.merge(lance_embutido, on="GRUPO", how="left")
    base = base.merge(lance_fidelidade, on="GRUPO", how="left")
    base = base.merge(tipos_por_grupo, on="GRUPO", how="left")
    base = base.merge(contemp_mes, on="GRUPO", how="left")
    base = base.merge(contemp_ll_mes, on="GRUPO", how="left")

    # Usar prazo das regras se não veio do histograma
    if "prazo_regras" in base.columns:
        base["prazo_original"] = base["prazo_original"].fillna(base["prazo_regras"])
        base = base.drop(columns=["prazo_regras"])

    # Preencher taxa adm padrão
    base["taxa_adm"] = TAXA_ADM_PADRAO

    # Remover grupos sem prazo
    base = base.dropna(subset=["prazo_original"])
    base["prazo_original"] = base["prazo_original"].astype(int)

    # Lance total médio em parcelas
    base["lance_embutido_parcelas"] = base["lance_embutido_parcelas"].fillna(0)
    base["lance_livre_parcelas"] = base["lance_livre_parcelas"].fillna(0)
    base["lance_fidelidade_parcelas"] = base["lance_fidelidade_parcelas"].fillna(0)
    base["lance_total_parcelas"] = (
        base["lance_livre_parcelas"]
        + base["lance_embutido_parcelas"]
        + base["lance_fidelidade_parcelas"]
    )

    # Filtrar grupos com pelo menos algum lance registrado
    base = base[base["lance_total_parcelas"] > 0].copy()
    base = base.fillna({"contemp_por_mes": 0, "contemp_ll_mes": 0})

    base = base.sort_values("GRUPO").reset_index(drop=True)
    return base


def _extrair_prazo_das_regras(df_regras: pd.DataFrame) -> pd.DataFrame:
    """
    Extrai mapeamento grupo → prazo da aba Regras.
    Colunas PRAZO 240, 220, 200, 180, 150, 100 nas posições 7-12.
    """
    prazo_cols = {7: 240, 8: 220, 9: 200, 10: 180, 11: 150, 12: 100}
    registros = []
    for col_idx, prazo in prazo_cols.items():
        if col_idx < df_regras.shape[1]:
            for val in df_regras.iloc[1:, col_idx].dropna():
                try:
                    g = int(float(val))
                    registros.append({"GRUPO": g, "prazo_regras": prazo})
                except (ValueError, TypeError):
                    pass
    if not registros:
        return pd.DataFrame(columns=["GRUPO", "prazo_regras"])
    return pd.DataFrame(registros).drop_duplicates("GRUPO")


def enriquecer_com_credito(
    base: pd.DataFrame, credito_map: dict
) -> pd.DataFrame:
    """
    Adiciona o crédito em R$ para cada grupo.
    credito_map: {grupo_int: valor_float}
    Usado quando o usuário fornece os valores manualmente.
    """
    base = base.copy()
    base["credito"] = base["GRUPO"].map(credito_map)
    return base
