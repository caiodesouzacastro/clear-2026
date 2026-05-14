"""
CLEAR 2026 — Dashboard de Organização e Alocação
FGV CLEAR
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import date
from pathlib import Path

st.set_page_config(
    page_title="CLEAR 2026",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

MASTER_FILE = Path(__file__).parent / "CLEAR_Master_2026.xlsx"

# Paleta Bloomberg dark
BG = "#0A1929"
BG_CARD = "#132F4C"
BG_HOVER = "#1E3A5F"
TEXTO = "#FFFFFF"
TEXTO_DIM = "#B2BAC2"
TEXTO_DIM2 = "#6F7E8C"
BORDA = "#1E3A5F"
AZUL = "#5090D3"
AZUL_CLARO = "#7BB3F0"
AZUL_ESCURO = "#1E3A5F"

CORES_STATUS = {
    "Concluído": "#6F7E8C",
    "Em Andamento": "#5090D3",
    "Não Iniciado": "#3A5169",
    "Atrasado": "#7BB3F0",
    "Reunião": "#4A6FA5",
}

st.markdown(f"""
<style>
    .stApp {{ background-color: {BG}; }}
    .main, .block-container {{ background-color: {BG}; color: {TEXTO}; }}
    [data-testid="stSidebar"] {{ background-color: {BG_CARD}; }}
    [data-testid="stSidebar"] > div {{ background-color: {BG_CARD}; }}
    h1, h2, h3, h4, h5, h6 {{ color: {TEXTO} !important; font-weight: 600; }}
    p, span, label, div {{ color: {TEXTO}; }}
    [data-testid="stMetricValue"] {{ color: {TEXTO} !important; font-weight: 700; font-size: 2.2rem; }}
    [data-testid="stMetricLabel"] {{ color: {TEXTO_DIM} !important; font-size: 0.85rem; text-transform: uppercase; letter-spacing: 0.05em; }}
    .stTabs [data-baseweb="tab-list"] {{ gap: 4px; background-color: {BG}; border-bottom: 1px solid {BORDA}; }}
    .stTabs [data-baseweb="tab"] {{ background-color: transparent; color: {TEXTO_DIM}; padding: 10px 20px; border-radius: 0; border-bottom: 2px solid transparent; }}
    .stTabs [aria-selected="true"] {{ background-color: transparent !important; color: {AZUL} !important; border-bottom: 2px solid {AZUL} !important; font-weight: 600; }}
    [data-baseweb="tag"] {{ background-color: {BG_HOVER} !important; color: {TEXTO} !important; }}
    [data-baseweb="select"] > div {{ background-color: {BG_CARD} !important; border-color: {BORDA} !important; }}
    [data-baseweb="input"] > div {{ background-color: {BG_CARD} !important; }}
    [data-testid="stDataFrame"] {{ background-color: {BG_CARD}; }}
    [data-testid="stExpander"] {{ background-color: {BG_CARD}; border: 1px solid {BORDA}; border-radius: 4px; }}
    hr {{ border-color: {BORDA}; }}
    header[data-testid="stHeader"] {{ background-color: {BG}; }}
</style>
""", unsafe_allow_html=True)


@st.cache_data(show_spinner="Carregando...")
def carregar_dados():
    atividades = pd.read_excel(MASTER_FILE, sheet_name="Atividades")
    resp = pd.read_excel(MASTER_FILE, sheet_name="Resp_Atividade")
    alocacao = pd.read_excel(MASTER_FILE, sheet_name="Alocacao")
    pessoas = pd.read_excel(MASTER_FILE, sheet_name="Pessoas")
    atividades["prazo"] = pd.to_datetime(atividades["prazo"], errors="coerce")
    resp["prazo"] = pd.to_datetime(resp["prazo"], errors="coerce")
    alocacao["data_mes"] = pd.to_datetime(alocacao["data_mes"], errors="coerce")
    atividades["eh_entregavel"] = atividades["eh_entregavel"].fillna(False).astype(bool)
    resp["eh_entregavel"] = resp["eh_entregavel"].fillna(False).astype(bool)
    atividades["responsaveis"] = atividades["responsaveis"].fillna("")
    return atividades, resp, alocacao, pessoas


def estilizar_fig(fig, altura=380):
    fig.update_layout(
        height=altura,
        margin=dict(t=30, b=30, l=10, r=10),
        paper_bgcolor=BG,
        plot_bgcolor=BG,
        font=dict(color=TEXTO_DIM, family="sans-serif"),
        xaxis=dict(gridcolor=BORDA, zerolinecolor=BORDA, color=TEXTO_DIM),
        yaxis=dict(gridcolor=BORDA, zerolinecolor=BORDA, color=TEXTO_DIM),
        legend=dict(bgcolor=BG_CARD, bordercolor=BORDA, font=dict(color=TEXTO)),
    )
    return fig


def aplicar_filtros_globais(df_at, df_resp):
    with st.sidebar:
        st.markdown("### CLEAR 2026")
        st.caption("Organização e Alocação")
        st.markdown("---")
        st.markdown("**Período**")
        hoje = pd.Timestamp(date.today())
        periodo = st.radio(
            "Janela",
            ["Próximas 4 semanas", "Próximos 3 meses", "Resto do ano", "Tudo"],
            index=1,
            label_visibility="collapsed",
        )
        if periodo == "Próximas 4 semanas":
            data_min, data_max = hoje, hoje + pd.Timedelta(days=28)
        elif periodo == "Próximos 3 meses":
            data_min, data_max = hoje, hoje + pd.Timedelta(days=90)
        elif periodo == "Resto do ano":
            data_min, data_max = hoje, pd.Timestamp("2026-12-31")
        else:
            data_min, data_max = pd.Timestamp("2026-01-01"), pd.Timestamp("2027-12-31")

        st.markdown("**Projetos**")
        projetos_disp = sorted(df_at["projeto"].dropna().unique())
        projs_selecionados = st.multiselect(
            "Selecione", options=projetos_disp, default=projetos_disp,
            label_visibility="collapsed",
        )
        st.markdown("**Status**")
        status_disp = sorted(df_at["status"].dropna().unique())
        status_selecionados = st.multiselect(
            "Selecione", options=status_disp,
            default=[s for s in status_disp if s not in ("Concluído", "Reunião")],
            label_visibility="collapsed",
        )
        st.markdown("---")
        st.caption(f"Hoje: {date.today().strftime('%d/%m/%Y')}")

    mask_at = df_at["projeto"].isin(projs_selecionados) & df_at["status"].isin(status_selecionados)
    mask_resp = df_resp["projeto"].isin(projs_selecionados) & df_resp["status"].isin(status_selecionados)
    return df_at[mask_at].copy(), df_resp[mask_resp].copy(), data_min, data_max


def view_portfolio(df_at, data_min, data_max):
    st.subheader("Portfólio dos Projetos")
    com_data = df_at[df_at["prazo"].notna()].copy()
    no_periodo = com_data[(com_data["prazo"] >= data_min) & (com_data["prazo"] <= data_max)]
    hoje_ts = pd.Timestamp(date.today())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Atividades no período", len(no_periodo))
    c2.metric("Entregáveis críticos", int(no_periodo["eh_entregavel"].sum()))
    c3.metric("Em andamento", int((no_periodo["status"] == "Em Andamento").sum()))
    atrasados = no_periodo[(no_periodo["status"] != "Concluído") & (no_periodo["prazo"] < hoje_ts)]
    c4.metric("Atrasados", len(atrasados))

    st.markdown("---")
    st.markdown("**Atividades por projeto no período**")
    if not no_periodo.empty:
        por_projeto = (
            no_periodo.groupby("projeto")
            .agg(total=("id_atividade", "count"), entregaveis=("eh_entregavel", "sum"))
            .reset_index()
            .sort_values("total", ascending=True)
        )
        fig = go.Figure()
        fig.add_trace(go.Bar(
            y=por_projeto["projeto"],
            x=por_projeto["total"] - por_projeto["entregaveis"],
            orientation="h", marker_color=AZUL_ESCURO,
            name="Atividades regulares",
            hovertemplate="<b>%{y}</b><br>Regulares: %{x}<extra></extra>",
        ))
        fig.add_trace(go.Bar(
            y=por_projeto["projeto"], x=por_projeto["entregaveis"],
            orientation="h", marker_color=AZUL,
            name="Entregáveis críticos",
            hovertemplate="<b>%{y}</b><br>Entregáveis: %{x}<extra></extra>",
        ))
        fig.update_layout(barmode="stack", xaxis_title="", yaxis_title="")
        fig = estilizar_fig(fig, altura=380)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    else:
        st.info("Sem atividades no período filtrado.")

    with st.expander("Ver timeline detalhada (atividade por atividade)"):
        if not no_periodo.empty:
            no_periodo_tl = no_periodo.copy()
            no_periodo_tl["tamanho"] = no_periodo_tl["eh_entregavel"].map({True: 16, False: 7})
            no_periodo_tl["atividade_curta"] = no_periodo_tl["atividade"].str[:90]
            fig = px.scatter(
                no_periodo_tl, x="prazo", y="projeto", color="status",
                color_discrete_map=CORES_STATUS, size="tamanho", size_max=16,
                hover_name="atividade_curta",
                hover_data={"responsaveis": True, "sub_projeto": True,
                            "prazo": "|%d/%m/%Y", "tamanho": False, "atividade_curta": False},
            )
            fig.add_shape(type="line", x0=hoje_ts, x1=hoje_ts, y0=0, y1=1, yref="paper",
                          line=dict(color=AZUL_CLARO, width=2, dash="dash"))
            fig.add_annotation(x=hoje_ts, y=1, yref="paper", text="hoje",
                               showarrow=False, font=dict(color=AZUL_CLARO), yshift=10)
            fig = estilizar_fig(fig, altura=460)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown("---")
    st.markdown("**Próximos entregáveis no período**")
    entregaveis = no_periodo[no_periodo["eh_entregavel"]].sort_values("prazo")
    if entregaveis.empty:
        st.info("Sem entregáveis no período filtrado.")
    else:
        st.dataframe(
            entregaveis[["prazo", "projeto", "sub_projeto", "atividade", "responsaveis", "status"]].rename(
                columns={"prazo": "Prazo", "projeto": "Projeto", "sub_projeto": "Subprojeto",
                         "atividade": "Entregável", "responsaveis": "Responsáveis", "status": "Status"}),
            use_container_width=True, hide_index=True,
            column_config={"Prazo": st.column_config.DateColumn(format="DD/MM/YYYY")},
        )


def view_pesquisador(df_resp, df_alocacao, df_pessoas, data_min, data_max):
    st.subheader("Pesquisador")
    pessoas_lista = df_pessoas["pessoa"].sort_values().tolist()
    pessoa = st.selectbox("Pesquisador(a)", pessoas_lista)

    aloc_pessoa = df_alocacao[df_alocacao["pessoa"] == pessoa].sort_values("data_mes").copy()
    if not aloc_pessoa.empty:
        carga_sem = aloc_pessoa["carga_horaria_semanal"].iloc[0]
        media_proj = aloc_pessoa["n_projetos_no_mes"].mean()
        max_proj = aloc_pessoa["n_projetos_no_mes"].max()
        pico_mes = aloc_pessoa.loc[aloc_pessoa["n_projetos_no_mes"].idxmax(), "mes"]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Dedicação", f"{carga_sem:.0f}h/sem")
        c2.metric("Média projetos/mês", f"{media_proj:.1f}")
        c3.metric("Pico de projetos", int(max_proj), help=f"em {pico_mes}")
        n_at = (df_resp[df_resp["responsavel"] == pessoa]["status"] != "Concluído").sum()
        c4.metric("Atividades em aberto", int(n_at))

        st.markdown("---")
        st.markdown("**Projetos simultâneos por mês**")

        def cor_n(n):
            if n >= 8: return AZUL_CLARO
            if n >= 6: return AZUL
            return AZUL_ESCURO

        cores = [cor_n(n) for n in aloc_pessoa["n_projetos_no_mes"]]
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=aloc_pessoa["mes"], y=aloc_pessoa["n_projetos_no_mes"],
            marker_color=cores, text=aloc_pessoa["n_projetos_no_mes"],
            textposition="outside", textfont=dict(color=TEXTO),
            hovertemplate="<b>%{x}</b><br>Projetos: %{y}<extra></extra>",
        ))
        fig.update_layout(yaxis_title="", xaxis_title="", showlegend=False)
        fig = estilizar_fig(fig, altura=320)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown("---")
    st.markdown(f"**Atividades de {pessoa}**")
    at_pessoa = df_resp[df_resp["responsavel"] == pessoa].copy()
    at_com_data = at_pessoa[at_pessoa["prazo"].notna()]

    if at_com_data.empty:
        st.info(f"{pessoa} não tem atividades com data atribuídas.")
    else:
        st.dataframe(
            at_com_data.sort_values("prazo")[
                ["prazo", "projeto", "sub_projeto", "atividade", "status", "eh_entregavel"]
            ].rename(columns={"prazo": "Prazo", "projeto": "Projeto", "sub_projeto": "Subprojeto",
                              "atividade": "Atividade", "status": "Status", "eh_entregavel": "Entregável?"}),
            use_container_width=True, hide_index=True,
            column_config={"Prazo": st.column_config.DateColumn(format="DD/MM/YYYY")},
        )


def view_proximas_entregas(df_at, data_min, data_max):
    st.subheader("Próximas Entregas")
    st.caption("Entregas críticas agrupadas por urgência.")

    entregaveis = df_at[df_at["eh_entregavel"] & df_at["prazo"].notna()].copy()
    entregaveis_periodo = entregaveis[
        (entregaveis["prazo"] >= data_min) & (entregaveis["prazo"] <= data_max)
    ].sort_values("prazo")

    if entregaveis_periodo.empty:
        st.info("Nenhum entregável crítico no período.")
        return

    hoje = pd.Timestamp(date.today())
    atrasados = entregaveis_periodo[(entregaveis_periodo["prazo"] < hoje) & (entregaveis_periodo["status"] != "Concluído")]
    proximos_7 = entregaveis_periodo[(entregaveis_periodo["prazo"] >= hoje) & (entregaveis_periodo["prazo"] <= hoje + pd.Timedelta(days=7))]
    proximos_30 = entregaveis_periodo[(entregaveis_periodo["prazo"] > hoje + pd.Timedelta(days=7)) & (entregaveis_periodo["prazo"] <= hoje + pd.Timedelta(days=30))]
    futuros = entregaveis_periodo[entregaveis_periodo["prazo"] > hoje + pd.Timedelta(days=30)]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Atrasados", len(atrasados))
    c2.metric("Próx. 7 dias", len(proximos_7))
    c3.metric("Próx. 30 dias", len(proximos_30))
    c4.metric("Depois", len(futuros))

    st.markdown("---")

    def mostrar_bloco(titulo, df, limite=15):
        if df.empty: return
        st.markdown(f"**{titulo}** ({len(df)})")
        df_mostrar = df.head(limite)
        for _, row in df_mostrar.iterrows():
            dias = (row["prazo"] - hoje).days
            if dias < 0:
                dias_txt = f"<b>{abs(dias)} dias atrasado</b>"
            elif dias == 0:
                dias_txt = "<b>hoje</b>"
            else:
                dias_txt = f"em {dias} dias"
            cor = CORES_STATUS.get(row["status"], TEXTO_DIM)
            resp_txt = row["responsaveis"] if row["responsaveis"] else "<i>sem responsável</i>"
            st.markdown(f"""
                <div style="border-left: 3px solid {cor}; padding: 10px 14px; margin-bottom: 8px;
                            background: {BG_CARD}; border-radius: 4px;">
                  <div style="font-size: 0.8rem; color: {TEXTO_DIM};">
                    <b style="color: {TEXTO};">{row['prazo'].strftime('%d/%m/%Y')}</b>
                    · {dias_txt} · <span style="color: {cor};">{row['status']}</span>
                  </div>
                  <div style="font-weight: 600; margin: 4px 0; color: {TEXTO};">{row['atividade']}</div>
                  <div style="font-size: 0.85rem; color: {TEXTO_DIM};">
                    <i>{row['projeto']}</i> — {row['sub_projeto']}<br>{resp_txt}
                  </div>
                </div>
            """, unsafe_allow_html=True)
        if len(df) > limite:
            st.caption(f"...e mais {len(df) - limite} item(ns).")
        st.markdown("")

    mostrar_bloco("Atrasados", atrasados)
    mostrar_bloco("Próximos 7 dias", proximos_7)
    mostrar_bloco("Próximos 30 dias", proximos_30)
    mostrar_bloco("Mais à frente", futuros)


def main():
    try:
        df_at, df_resp, df_alocacao, df_pessoas = carregar_dados()
    except FileNotFoundError:
        st.error("Planilha mestra não encontrada.")
        st.stop()

    df_at_f, df_resp_f, data_min, data_max = aplicar_filtros_globais(df_at, df_resp)

    tab1, tab2, tab3 = st.tabs(["Portfólio", "Pesquisador", "Próximas Entregas"])
    with tab1:
        view_portfolio(df_at_f, data_min, data_max)
    with tab2:
        view_pesquisador(df_resp_f, df_alocacao, df_pessoas, data_min, data_max)
    with tab3:
        view_proximas_entregas(df_at_f, data_min, data_max)


if __name__ == "__main__":
    main()
