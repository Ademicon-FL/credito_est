"""
engine.py
Motor de cálculo financeiro por grupo.
Suporta FIDC% configurável (0 = sem FIDC, cliente usa capital próprio).
"""

import numpy as np
from financial import irr as npf_irr
import pandas as pd
from dataclasses import dataclass
from typing import Optional


@dataclass
class ResultadoGrupo:
    grupo: int
    credito: float
    prazo_original: int
    prazo_atual: int
    taxa_adm: float
    participantes: int

    # Lance
    lance_total_parcelas: float
    lance_embutido_parcelas: float
    lance_livre_parcelas: float
    lance_R: float              # Lance total em R$
    lance_embutido_R: float     # Parcela embutida no crédito
    lance_livre_R: float        # Parcela que precisa de capital (FIDC ou próprio)
    lance_pct: float            # % do crédito

    # Parcelas
    parcela_lancamento: float   # Parcela mensal (ao comprar)
    nova_parcela: float         # Parcela pós contemplação

    # Crédito
    ss_novo: float              # Crédito bruto liberado (sem FIDC)
    fidc_fee: float             # Taxa FIDC em R$ (0 se fidc_pct=0)
    credito_liquido: float      # Crédito líquido final ao cliente
    fidc_pct: float             # % FIDC aplicado

    # Custo
    custo_mensal: float         # TIR mensal
    custo_anual: float          # TIR anual equivalente
    meses_ate_contemplacao: int # Prazo_orig - prazo_atual (estimativa)

    tipos_lance: str            # Tipos disponíveis neste grupo
    contemp_ll_mes: float       # Média de contemplações LL por mês


def calcular_grupo(
    grupo: int,
    credito: float,
    prazo_original: int,
    prazo_atual: int,
    taxa_adm: float,
    participantes: int,
    lance_total_parcelas: float,
    lance_embutido_parcelas: float,
    lance_livre_parcelas: float,
    fidc_pct: float = 0.0,
    tipos_lance: str = "Lance Livre",
    contemp_ll_mes: float = 0,
) -> Optional[ResultadoGrupo]:
    """
    Calcula todos os indicadores financeiros para um grupo/cota.

    Parâmetros principais:
      credito:               Valor do crédito em R$
      prazo_original:        Prazo original do consórcio em meses
      prazo_atual:           Prazo restante atual em meses
      taxa_adm:              Taxa de administração total (ex: 0.24 = 24%)
      lance_total_parcelas:  Média de parcelas ofertadas no lance
      lance_embutido_pct:    Parcelas do lance que vêm do próprio crédito (fixo/limitado)
      fidc_pct:              % de taxa do FIDC (0 = sem FIDC, 0.05 = 5%)

    Fórmulas verificadas contra planilha PrevOne real:
      parcela_lancamento = credito × (1 + taxa_adm) / prazo_original
      lance_R            = lance_total_parcelas × parcela_lancamento
      ss_novo            = credito - lance_R
      nova_parcela       = (ss_novo + credito × taxa_adm) / prazo_atual
      fidc_fee           = lance_livre_R × fidc_pct
      credito_liquido    = ss_novo - fidc_fee
    """
    if credito <= 0 or prazo_original <= 0 or prazo_atual <= 0:
        return None
    if lance_total_parcelas <= 0:
        return None

    # Parcela mensal ao comprar a cota
    parcela_lancamento = credito * (1 + taxa_adm) / prazo_original

    # Lance em R$
    lance_R = lance_total_parcelas * parcela_lancamento
    lance_embutido_R = lance_embutido_parcelas * parcela_lancamento
    lance_livre_R = lance_R - lance_embutido_R

    # Proteção: lance não pode superar o crédito
    if lance_R >= credito:
        return None

    # Crédito bruto liberado ao cliente
    ss_novo = credito - lance_R
    lance_pct = lance_R / credito

    # Parcela pós contemplação (fórmula confirmada)
    nova_parcela = (ss_novo + credito * taxa_adm) / prazo_atual

    # FIDC: cobra taxa sobre o lance livre financiado
    # Se fidc_pct = 0: cliente usa capital próprio, sem custo adicional de FIDC
    fidc_fee = lance_livre_R * fidc_pct
    credito_liquido = ss_novo - fidc_fee

    if credito_liquido <= 0:
        return None

    # Meses estimados até contemplação
    meses_ate = max(1, int(prazo_original - prazo_atual))

    # TIR: fluxos de caixa da perspectiva do cliente
    #   Pré-contemplação: paga parcela_lancamento por meses_ate meses
    #   Contemplação:     recebe credito_liquido (já descontado FIDC)
    #   Pós-contemplação: paga nova_parcela por prazo_atual meses
    cfs = (
        [-parcela_lancamento] * meses_ate
        + [credito_liquido]
        + [-nova_parcela] * int(prazo_atual)
    )

    # CUSTO = IRR de [ss_novo, -nova_parcela × prazo_atual]
    # Fórmula confirmada contra todos os grupos da planilha PrevOne real
    cfs_custo = [ss_novo] + [-nova_parcela] * int(prazo_atual)
    try:
        custo_mensal = npf_irr(cfs_custo)
        if np.isnan(custo_mensal) or custo_mensal <= 0:
            custo_mensal = 0.01
    except Exception:
        custo_mensal = 0.01

    custo_anual = (1 + custo_mensal) ** 12 - 1

    return ResultadoGrupo(
        grupo=grupo,
        credito=credito,
        prazo_original=prazo_original,
        prazo_atual=prazo_atual,
        taxa_adm=taxa_adm,
        participantes=participantes,
        lance_total_parcelas=lance_total_parcelas,
        lance_embutido_parcelas=lance_embutido_parcelas,
        lance_livre_parcelas=lance_livre_parcelas,
        lance_R=lance_R,
        lance_embutido_R=lance_embutido_R,
        lance_livre_R=lance_livre_R,
        lance_pct=lance_pct,
        parcela_lancamento=parcela_lancamento,
        nova_parcela=nova_parcela,
        ss_novo=ss_novo,
        fidc_fee=fidc_fee,
        credito_liquido=credito_liquido,
        fidc_pct=fidc_pct,
        custo_mensal=custo_mensal,
        custo_anual=custo_anual,
        meses_ate_contemplacao=meses_ate,
        tipos_lance=tipos_lance,
        contemp_ll_mes=contemp_ll_mes,
    )


def processar_base_grupos(
    df_base: pd.DataFrame,
    fidc_pct: float = 0.0,
    credito_map: dict = None,
) -> pd.DataFrame:
    """
    Processa a base de grupos e retorna DataFrame com todos os indicadores.

    df_base: saída do extractor.py (uma linha por grupo)
    fidc_pct: % FIDC aplicado (0 a 1)
    credito_map: dict {grupo: valor_credito} para grupos onde o usuário informou o valor
    """
    resultados = []

    for _, row in df_base.iterrows():
        grupo = int(row["GRUPO"])

        # Crédito: do mapa do usuário ou da planilha
        if credito_map and grupo in credito_map:
            credito = float(credito_map[grupo])
        elif "credito" in row and pd.notna(row.get("credito", None)):
            credito = float(row["credito"])
        else:
            continue  # sem crédito, skip

        resultado = calcular_grupo(
            grupo=grupo,
            credito=credito,
            prazo_original=int(row["prazo_original"]),
            prazo_atual=int(row.get("prazo_atual", row["prazo_original"])),
            taxa_adm=float(row.get("taxa_adm", 0.24)),
            participantes=int(row.get("participantes", 0)),
            lance_total_parcelas=float(row.get("lance_total_parcelas", 0)),
            lance_embutido_parcelas=float(row.get("lance_embutido_parcelas", 0)),
            lance_livre_parcelas=float(row.get("lance_livre_parcelas", 0)),
            fidc_pct=fidc_pct,
            tipos_lance=str(row.get("tipos_lance", "")),
            contemp_ll_mes=float(row.get("contemp_ll_mes", 0)),
        )

        if resultado:
            resultados.append(resultado.__dict__)

    if not resultados:
        return pd.DataFrame()

    df_resultado = pd.DataFrame(resultados)
    df_resultado = df_resultado.sort_values("custo_mensal").reset_index(drop=True)
    return df_resultado
