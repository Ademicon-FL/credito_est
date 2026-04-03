"""
optimizer.py
Algoritmo de otimização da carteira de cotas para atingir crédito alvo.
Distribui as contemplações ao longo de N meses conforme solicitado.
"""

import pandas as pd
import numpy as np
from dataclasses import dataclass
from typing import List, Tuple
from financial import irr as npf_irr


@dataclass
class LinhaCarteira:
    grupo: int
    credito: float
    prazo_original: int
    prazo_atual: int
    taxa_adm: float
    parcela_lancamento: float
    lance_R: float
    lance_embutido_R: float
    lance_livre_R: float
    lance_pct: float
    qtde_cotas: int
    ss_novo_por_cota: float
    ss_novo_total: float
    fidc_fee_por_cota: float
    fidc_fee_total: float
    credito_liquido_por_cota: float
    credito_liquido_total: float
    nova_parcela: float
    custo_mensal: float
    tranche: int  # mês de contemplação (1, 2, 3...)
    tipos_lance: str


@dataclass
class ResultadoOtimizacao:
    carteira: List[LinhaCarteira]
    credito_liquido_total: float
    credito_bruto_total: float
    parcela_pre_total: float    # Parcela pré-contemplação (todas as cotas)
    parcela_pos_total: float    # Parcela final (todas pós contemplação)
    tir_mensal: float
    tir_anual: float
    num_cotas: int
    distribuicao_meses: int
    fidc_pct: float
    fidc_fee_total: float


def otimizar_carteira(
    df_grupos: pd.DataFrame,
    credito_alvo: float,
    fidc_pct: float = 0.0,
    meses_distribuicao: int = 4,
    max_cotas_por_grupo: int = 5,
    credito_minimo_grupo: float = 20000,
    apenas_lance_livre: bool = False,
) -> ResultadoOtimizacao:
    """
    Seleciona a combinação ótima de grupos e cotas para atingir o crédito alvo.

    Estratégia greedy:
      1. Ordena grupos por menor custo (TIR) e maior contemplação por mês
      2. Seleciona grupos com contemplação mensal ≥ qtde_cotas para cada grupo
      3. Distribui as contemplações ao longo dos meses de distribuição
      4. Para quando atingir o crédito alvo

    Parâmetros:
      credito_alvo:       Crédito líquido desejado pelo cliente (R$)
      fidc_pct:           % de taxa do FIDC (0 = sem FIDC)
      meses_distribuicao: Em quantos meses distribuir as contemplações (ex: 4)
      max_cotas_por_grupo: Máximo de cotas do mesmo grupo na carteira
      apenas_lance_livre:  Se True, exclui grupos com lance embutido
    """
    df = df_grupos.copy()

    # Filtros de elegibilidade
    df = df[df["credito_liquido"] >= credito_minimo_grupo]
    if apenas_lance_livre:
        df = df[df["lance_embutido_R"] == 0]

    if df.empty:
        raise ValueError("Nenhum grupo elegível com os filtros aplicados.")

    # Score de prioridade: menor custo + maior frequência de contemplação
    df["score"] = (
        -df["custo_mensal"] * 0.6
        + (df["contemp_ll_mes"] / (df["contemp_ll_mes"].max() + 1e-9)) * 0.4
    )
    df = df.sort_values("score", ascending=False).reset_index(drop=True)

    # Construção gulosa da carteira
    carteira: List[LinhaCarteira] = []
    credito_acumulado = 0.0
    tranche = 1

    for _, row in df.iterrows():
        if credito_acumulado >= credito_alvo:
            break

        # Quantas cotas precisamos ainda?
        faltando = credito_alvo - credito_acumulado
        cota_liquido = row["credito_liquido"]

        # Cotas necessárias (limitado pelo max permitido)
        cotas_necessarias = max(1, int(np.ceil(faltando / cota_liquido)))
        cotas_necessarias = min(cotas_necessarias, max_cotas_por_grupo)

        # Distribuição por tranche
        tranche_atual = ((len(carteira)) % meses_distribuicao) + 1

        linha = LinhaCarteira(
            grupo=int(row["grupo"]),
            credito=row["credito"],
            prazo_original=int(row["prazo_original"]),
            prazo_atual=int(row["prazo_atual"]),
            taxa_adm=row["taxa_adm"],
            parcela_lancamento=row["parcela_lancamento"],
            lance_R=row["lance_R"],
            lance_embutido_R=row["lance_embutido_R"],
            lance_livre_R=row["lance_livre_R"],
            lance_pct=row["lance_pct"],
            qtde_cotas=cotas_necessarias,
            ss_novo_por_cota=row["ss_novo"],
            ss_novo_total=row["ss_novo"] * cotas_necessarias,
            fidc_fee_por_cota=row["fidc_fee"],
            fidc_fee_total=row["fidc_fee"] * cotas_necessarias,
            credito_liquido_por_cota=row["credito_liquido"],
            credito_liquido_total=row["credito_liquido"] * cotas_necessarias,
            nova_parcela=row["nova_parcela"],
            custo_mensal=row["custo_mensal"],
            tranche=tranche_atual,
            tipos_lance=row["tipos_lance"],
        )

        carteira.append(linha)
        credito_acumulado += linha.credito_liquido_total

    if not carteira:
        raise ValueError("Não foi possível montar carteira com os grupos disponíveis.")

    # Totais consolidados
    credito_liquido_total = sum(l.credito_liquido_total for l in carteira)
    credito_bruto_total = sum(l.ss_novo_total for l in carteira)
    fidc_fee_total = sum(l.fidc_fee_total for l in carteira)
    num_cotas = sum(l.qtde_cotas for l in carteira)

    parcela_pre_total = sum(
        l.parcela_lancamento * l.qtde_cotas for l in carteira
    )
    parcela_pos_total = sum(
        l.nova_parcela * l.qtde_cotas for l in carteira
    )

    # TIR agregada da operação completa
    tir_mensal, tir_anual = calcular_tir_operacao(carteira, meses_distribuicao)

    return ResultadoOtimizacao(
        carteira=carteira,
        credito_liquido_total=credito_liquido_total,
        credito_bruto_total=credito_bruto_total,
        parcela_pre_total=parcela_pre_total,
        parcela_pos_total=parcela_pos_total,
        tir_mensal=tir_mensal,
        tir_anual=tir_anual,
        num_cotas=num_cotas,
        distribuicao_meses=meses_distribuicao,
        fidc_pct=fidc_pct,
        fidc_fee_total=fidc_fee_total,
    )


def calcular_tir_operacao(
    carteira: List[LinhaCarteira],
    meses_distribuicao: int,
) -> Tuple[float, float]:
    """
    Calcula a TIR mensal e anual da operação completa.
    Modela o fluxo de caixa mês a mês conforme as contemplações ocorrem.
    """
    if not carteira:
        return 0.0, 0.0

    prazo_total = max(l.prazo_atual for l in carteira) + meses_distribuicao + 2

    # Fluxo de caixa por mês (perspectiva do cliente)
    cfs = [0.0] * prazo_total

    for linha in carteira:
        t = linha.tranche  # mês de contemplação (1-indexed)

        # Pré-contemplação: paga parcela desde o início até contemplação
        for m in range(0, t):
            if m < prazo_total:
                cfs[m] -= linha.parcela_lancamento * linha.qtde_cotas

        # Mês da contemplação: recebe o crédito líquido
        if t < prazo_total:
            cfs[t] += linha.credito_liquido_total
            cfs[t] -= linha.parcela_lancamento * linha.qtde_cotas  # ainda paga parcela neste mês

        # Pós-contemplação: paga nova_parcela
        for m in range(t + 1, min(t + int(linha.prazo_atual) + 1, prazo_total)):
            cfs[m] -= linha.nova_parcela * linha.qtde_cotas

    # Remove zeros do final
    while cfs and cfs[-1] == 0:
        cfs.pop()

    if len(cfs) < 2:
        return 0.01, (1.01**12 - 1)

    try:
        tir_mensal = npf_irr(cfs)
        if np.isnan(tir_mensal) or tir_mensal is None or tir_mensal < 0:
            tir_mensal = 0.01
    except Exception:
        tir_mensal = 0.01

    tir_anual = (1 + tir_mensal) ** 12 - 1
    return float(tir_mensal), float(tir_anual)


def gerar_fluxo_mensal(
    carteira: List[LinhaCarteira],
    meses_distribuicao: int,
) -> pd.DataFrame:
    """
    Gera o fluxo mensal de contemplações e créditos (aba FLUXO do PrevOne).
    """
    prazo_maximo = max(l.prazo_atual for l in carteira)
    total_meses = meses_distribuicao + prazo_maximo + 2

    linhas = []
    credito_acumulado = 0.0
    parcela_total_atual = sum(l.parcela_lancamento * l.qtde_cotas for l in carteira)

    for mes in range(1, total_meses + 1):
        # Cotas contempladas neste mês
        cotas_mes = [l for l in carteira if l.tranche == mes]
        cotas_contemp = sum(l.qtde_cotas for l in cotas_mes)

        credito_mes = sum(l.credito_liquido_total for l in cotas_mes)
        credito_acumulado += credito_mes

        # Parcela que o cliente paga neste mês
        # = soma das parcelas pré de quem ainda não contemplou
        # + soma das nova_parcela de quem já contemplou
        parcela_paga = 0.0
        for linha in carteira:
            if mes <= linha.tranche:
                parcela_paga += linha.parcela_lancamento * linha.qtde_cotas
            else:
                parcela_paga += linha.nova_parcela * linha.qtde_cotas

        # Caixa = crédito recebido - parcela paga
        caixa = credito_mes - parcela_paga

        linhas.append({
            "Mês": mes,
            "Cotas contempladas": cotas_contemp if cotas_contemp > 0 else None,
            "Parcela paga (R$)": parcela_paga,
            "Crédito liberado (R$)": credito_mes if credito_mes > 0 else None,
            "Crédito acumulado (R$)": credito_acumulado if credito_acumulado > 0 else None,
            "Caixa (R$)": caixa,
        })

        # Para quando não houver mais parcelas relevantes
        if mes > meses_distribuicao + 3 and credito_mes == 0:
            ultimo_prazo = min(l.prazo_atual for l in carteira)
            if mes > meses_distribuicao + ultimo_prazo:
                break

    return pd.DataFrame(linhas)
