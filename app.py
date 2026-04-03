"""
app.py
Interface Streamlit do gerador PrevOne de crédito estruturado via consórcio.
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io

# ─── Configuração da página ───────────────────────────────────────────────────
st.set_page_config(
    page_title="PrevOne · Crédito Estruturado",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS minimalista ──────────────────────────────────────────────────────────
st.markdown("""
<style>
    .metric-card {
        background: #f0f4ff;
        border-left: 4px solid #1F3864;
        border-radius: 6px;
        padding: 12px 16px;
        margin-bottom: 8px;
    }
    .metric-label { font-size: 12px; color: #555; font-weight: 500; }
    .metric-value { font-size: 22px; color: #1F3864; font-weight: 700; }
    .metric-sub   { font-size: 11px; color: #888; }
    .section-title {
        font-size: 15px; font-weight: 700; color: #1F3864;
        border-bottom: 2px solid #1F3864;
        padding-bottom: 4px; margin: 16px 0 10px 0;
    }
    .tag-fidc { background:#e8f4e8; color:#2d6a2d; padding:2px 8px; border-radius:12px; font-size:12px; font-weight:600; }
    .tag-proprio { background:#fff3e0; color:#b26a00; padding:2px 8px; border-radius:12px; font-size:12px; font-weight:600; }
    .aviso { font-size:11px; color:#888; font-style:italic; margin-top:8px; }
</style>
""", unsafe_allow_html=True)


# ─── Importações dos módulos locais ──────────────────────────────────────────
try:
    from extractor import extrair_base_grupos, enriquecer_com_credito
    from engine import processar_base_grupos
    from optimizer import otimizar_carteira
    from generator import gerar_excel
    MODULOS_OK = True
except ImportError as e:
    MODULOS_OK = False
    st.error(f"Erro ao importar módulos: {e}")


# ─── Sidebar: Parâmetros da Operação ─────────────────────────────────────────
with st.sidebar:
    st.image("https://via.placeholder.com/200x50/1F3864/FFFFFF?text=PrevOne", width=200)
    st.markdown("### Parâmetros da operação")

    nome_cliente = st.text_input("Nome do cliente", value="", placeholder="Ex: João Silva")

    st.markdown("---")
    st.markdown("**Crédito e distribuição**")

    credito_alvo = st.number_input(
        "Crédito líquido alvo (R$)",
        min_value=50_000.0,
        max_value=20_000_000.0,
        value=1_500_000.0,
        step=50_000.0,
        format="%.0f",
        help="Valor total líquido que o cliente deseja levantar",
    )

    meses_distribuicao = st.slider(
        "Meses de distribuição",
        min_value=1, max_value=12, value=4,
        help="Em quantos meses as contemplações serão distribuídas",
    )

    max_cotas_por_grupo = st.slider(
        "Máx. cotas por grupo",
        min_value=1, max_value=10, value=3,
        help="Limite de cotas do mesmo grupo na carteira",
    )

    st.markdown("---")
    st.markdown("**Financiamento FIDC**")

    usar_fidc = st.toggle("Usar FIDC para o lance", value=True)

    if usar_fidc:
        fidc_pct = st.slider(
            "Taxa FIDC (%)",
            min_value=1, max_value=15, value=5,
            help="% cobrado pelo FIDC sobre o lance livre financiado",
        ) / 100.0
        st.markdown(
            '<span class="tag-fidc">✓ FIDC cobre o lance</span> '
            '— cliente não precisa de capital próprio na contemplação.',
            unsafe_allow_html=True,
        )
    else:
        fidc_pct = 0.0
        st.markdown(
            '<span class="tag-proprio">$ Capital próprio</span> '
            '— cliente paga o lance com recursos próprios na contemplação.',
            unsafe_allow_html=True,
        )

    st.markdown("---")
    st.markdown("**Filtros de grupos**")

    meses_historico = st.slider(
        "Histórico de lances (meses)",
        min_value=1, max_value=12, value=6,
        help="Quantos meses retroativos usar para calcular o lance médio",
    )

    apenas_ll = st.checkbox(
        "Apenas Lance Livre",
        value=False,
        help="Excluir grupos com lance fixo/limitado (embutido)",
    )

    st.markdown("---")
    data_base = st.date_input("Data base", value=datetime.today())


# ─── Conteúdo principal ───────────────────────────────────────────────────────
st.title("🏠 PrevOne · Gerador de Crédito Estruturado")
st.markdown(
    "Upload do histograma mensal da Ademicon → preenchimento dos créditos → "
    "download da planilha PrevOne."
)

# ─── Passo 1: Upload do histograma ───────────────────────────────────────────
st.markdown('<div class="section-title">1 · Histograma Mensal (Ademicon)</div>', unsafe_allow_html=True)

col_upload, col_info = st.columns([2, 1])

with col_upload:
    hist_file = st.file_uploader(
        "Arraste ou selecione o arquivo .xlsx",
        type=["xlsx"],
        key="histograma",
        help="Arquivo exportado mensalmente da plataforma Ademicon (Imóveis)",
    )

with col_info:
    st.info(
        "📋 **O app extrai automaticamente:**\n"
        "- Lance médio por grupo (últimos N meses)\n"
        "- Frequência de contemplações LL\n"
        "- Prazo original de cada grupo\n\n"
        "Você informa apenas o **crédito em R$** por grupo."
    )


# ─── Passo 2: Base de grupos ──────────────────────────────────────────────────
if hist_file and MODULOS_OK:
    with st.spinner("Lendo histograma..."):
        try:
            import tempfile, os
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(hist_file.read())
                tmp_path = tmp.name

            df_extraido = extrair_base_grupos(tmp_path, meses_historico=meses_historico)
            os.unlink(tmp_path)

            if df_extraido.empty:
                st.error("Não foram encontrados grupos com lance no histograma.")
                st.stop()

            st.success(f"✓ {len(df_extraido)} grupos extraídos do histograma.")

        except Exception as e:
            st.error(f"Erro ao ler histograma: {e}")
            st.stop()

    # ─── Passo 2: Tabela de créditos (editável) ───────────────────────────────
    st.markdown('<div class="section-title">2 · Crédito por grupo (preencha)</div>', unsafe_allow_html=True)

    st.markdown(
        "Informe o **crédito em R$** para os grupos de interesse. "
        "Deixe **0** para excluir o grupo da seleção."
    )

    # Montar tabela editável
    df_edit = df_extraido[["GRUPO", "prazo_original", "participantes",
                            "lance_total_parcelas", "lance_embutido_parcelas",
                            "contemp_ll_mes", "tipos_lance"]].copy()
    df_edit = df_edit.rename(columns={
        "GRUPO": "Grupo",
        "prazo_original": "Prazo (meses)",
        "participantes": "Participantes",
        "lance_total_parcelas": "Lance médio (parcelas)",
        "lance_embutido_parcelas": "Lance embutido (parcelas)",
        "contemp_ll_mes": "Contempl./mês LL",
        "tipos_lance": "Tipos de lance",
    })
    df_edit["Crédito (R$)"] = 0.0

    # Prefill de créditos a partir da sessão anterior
    if "creditos_salvos" in st.session_state:
        for i, row in df_edit.iterrows():
            g = int(row["Grupo"])
            if g in st.session_state["creditos_salvos"]:
                df_edit.at[i, "Crédito (R$)"] = st.session_state["creditos_salvos"][g]

    df_editado = st.data_editor(
        df_edit,
        column_config={
            "Grupo": st.column_config.NumberColumn("Grupo", disabled=True, format="%d"),
            "Prazo (meses)": st.column_config.NumberColumn("Prazo", disabled=True, format="%d"),
            "Participantes": st.column_config.NumberColumn("Part.", disabled=True, format="%d"),
            "Lance médio (parcelas)": st.column_config.NumberColumn(
                "Lance médio (parc.)", disabled=True, format="%.1f"
            ),
            "Lance embutido (parcelas)": st.column_config.NumberColumn(
                "Lance embutido (parc.)", disabled=True, format="%.1f"
            ),
            "Contempl./mês LL": st.column_config.NumberColumn(
                "Contempl./mês", disabled=True, format="%.1f"
            ),
            "Tipos de lance": st.column_config.TextColumn("Tipos", disabled=True),
            "Crédito (R$)": st.column_config.NumberColumn(
                "Crédito (R$) ✏️",
                min_value=0,
                format="R$ %.2f",
                help="Informe o valor do crédito para este grupo",
            ),
        },
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        key="tabela_grupos",
    )

    # Salvar créditos na sessão para persistência
    credito_map = {}
    for _, row in df_editado.iterrows():
        g = int(row["Grupo"])
        val = float(row["Crédito (R$)"])
        if val > 0:
            credito_map[g] = val
    st.session_state["creditos_salvos"] = credito_map

    grupos_com_credito = len(credito_map)
    st.caption(f"✓ {grupos_com_credito} grupos com crédito informado.")

    if grupos_com_credito == 0:
        st.warning("⚠️ Preencha o crédito (R$) de pelo menos um grupo para continuar.")
        st.stop()

    # ─── Passo 3: Calcular e gerar ────────────────────────────────────────────
    st.markdown('<div class="section-title">3 · Gerar operação</div>', unsafe_allow_html=True)

    col_btn, col_status = st.columns([1, 3])
    with col_btn:
        gerar = st.button("⚡ Gerar operação", type="primary", use_container_width=True)

    if gerar or "resultado" in st.session_state:

        if gerar:
            # Recalcula
            with st.spinner("Calculando operação..."):
                try:
                    # Enriquece com créditos
                    df_enriquecido = df_extraido.copy()
                    df_enriquecido["credito"] = df_enriquecido["GRUPO"].map(credito_map)

                    # Prazo atual = prazo original (histograma não tem prazo atual; usuário pode ajustar)
                    if "prazo_atual" not in df_enriquecido.columns:
                        df_enriquecido["prazo_atual"] = df_enriquecido["prazo_original"]

                    # Calcular financeiros por grupo
                    df_grupos_calc = processar_base_grupos(
                        df_enriquecido,
                        fidc_pct=fidc_pct,
                    )

                    if df_grupos_calc.empty:
                        st.error("Nenhum grupo elegível. Verifique os créditos informados.")
                        st.stop()

                    # Otimizar carteira
                    resultado = otimizar_carteira(
                        df_grupos_calc,
                        credito_alvo=credito_alvo,
                        fidc_pct=fidc_pct,
                        meses_distribuicao=meses_distribuicao,
                        max_cotas_por_grupo=max_cotas_por_grupo,
                        apenas_lance_livre=apenas_ll,
                    )

                    st.session_state["resultado"] = resultado
                    st.session_state["df_grupos_calc"] = df_grupos_calc

                except Exception as e:
                    st.error(f"Erro ao calcular: {e}")
                    import traceback
                    st.code(traceback.format_exc())
                    st.stop()

        resultado = st.session_state.get("resultado")
        df_grupos_calc = st.session_state.get("df_grupos_calc", pd.DataFrame())

        if resultado:
            # ─── Painel de resultados ─────────────────────────────────────────
            st.markdown("#### Resumo da operação")

            c1, c2, c3, c4 = st.columns(4)

            def metric_card(col, label, value, sub=""):
                col.markdown(
                    f'<div class="metric-card">'
                    f'<div class="metric-label">{label}</div>'
                    f'<div class="metric-value">{value}</div>'
                    f'<div class="metric-sub">{sub}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

            metric_card(c1, "Crédito líquido total",
                        f"R$ {resultado.credito_liquido_total:,.0f}",
                        f"Alvo: R$ {credito_alvo:,.0f}")
            metric_card(c2, "TIR anual da operação",
                        f"{resultado.tir_anual:.2%}",
                        f"TIR mensal: {resultado.tir_mensal:.4%}")
            metric_card(c3, "Parcela final (total)",
                        f"R$ {resultado.parcela_pos_total:,.0f}/mês",
                        f"Pré-contemplação: R$ {resultado.parcela_pre_total:,.0f}/mês")

            if fidc_pct > 0:
                metric_card(c4, f"Fee FIDC ({fidc_pct:.0%})",
                            f"R$ {resultado.fidc_fee_total:,.0f}",
                            f"{resultado.num_cotas} cotas · {meses_distribuicao} meses")
            else:
                metric_card(c4, "Total de cotas",
                            f"{resultado.num_cotas} cotas",
                            f"Em {meses_distribuicao} tranches · Sem FIDC")

            # ─── Tabela da carteira ────────────────────────────────────────────
            st.markdown("#### Carteira selecionada")

            dados_cart = []
            for l in resultado.carteira:
                dados_cart.append({
                    "Grupo": l.grupo,
                    "Tranche": l.tranche,
                    "Cotas": l.qtde_cotas,
                    "Crédito (R$)": l.credito,
                    "Lance (R$)": l.lance_R,
                    "Lance %": f"{l.lance_pct:.1%}",
                    "FIDC (R$)": l.fidc_fee_por_cota if fidc_pct > 0 else "—",
                    "Créd. Líquido/cota (R$)": l.credito_liquido_por_cota,
                    "Créd. Líquido Total (R$)": l.credito_liquido_total,
                    "Nova Parcela (R$)": l.nova_parcela,
                    "TIR mensal": f"{l.custo_mensal:.4%}",
                    "Tipo lance": l.tipos_lance,
                })

            df_cart_show = pd.DataFrame(dados_cart)
            st.dataframe(
                df_cart_show,
                hide_index=True,
                use_container_width=True,
                column_config={
                    "Crédito (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
                    "Lance (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
                    "FIDC (R$)": st.column_config.NumberColumn(format="R$ %.2f") if fidc_pct > 0 else None,
                    "Créd. Líquido/cota (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
                    "Créd. Líquido Total (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
                    "Nova Parcela (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
                },
            )

            # ─── Download ──────────────────────────────────────────────────────
            st.markdown("---")
            st.markdown("#### Download da planilha PrevOne")

            col_dl1, col_dl2 = st.columns([1, 3])

            with col_dl1:
                with st.spinner("Gerando Excel..."):
                    try:
                        excel_bytes = gerar_excel(
                            resultado=resultado,
                            df_base_grupos=df_grupos_calc,
                            nome_cliente=nome_cliente or "Cliente",
                            data_base=datetime.combine(data_base, datetime.min.time()),
                        )

                        nome_arquivo = (
                            f"PrevOne_{nome_cliente.replace(' ', '_') or 'Operacao'}"
                            f"_{datetime.today().strftime('%Y%m%d')}.xlsx"
                        )

                        st.download_button(
                            label="📥 Baixar planilha PrevOne",
                            data=excel_bytes,
                            file_name=nome_arquivo,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"Erro ao gerar Excel: {e}")
                        import traceback
                        st.code(traceback.format_exc())

            with col_dl2:
                st.markdown(
                    "A planilha contém 4 abas: **RESUMO** · **CARTEIRA** · **FLUXO** · **BASE GRUPOS**"
                )
                st.markdown(
                    '<p class="aviso">Simulação para fins ilustrativos. '
                    'Não configura promessa ou garantia de contemplação. '
                    'O percentual de lance reflete a média histórica de lances confirmados.</p>',
                    unsafe_allow_html=True,
                )

elif not hist_file:
    st.info("⬆️ Faça o upload do histograma mensal da Ademicon para começar.")

    # Modo demo: mostra estrutura esperada
    with st.expander("Ver estrutura esperada da planilha PrevOne gerada"):
        st.markdown("""
        **Aba RESUMO** — indicadores consolidados da operação:
        crédito líquido, TIR, parcelas pré/pós, fee FIDC, fluxo de liberação por mês.

        **Aba CARTEIRA** — detalhe de cada grupo selecionado:
        crédito, lance, embutido/livre, crédito líquido por cota, nova parcela, TIR individual.

        **Aba FLUXO** — fluxo mensal mês a mês:
        cotas contempladas, parcela paga, crédito liberado, caixa acumulado, TIR.

        **Aba BASE GRUPOS** — dados de todos os grupos analisados:
        lance médio, crédito, prazo, taxa adm, custo individual.
        """)
