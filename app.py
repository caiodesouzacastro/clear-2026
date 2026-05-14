"""
CLEAR 2026 — Dashboard de Organização e Alocação
FGV CLEAR

Rodar localmente:
    pip install -r requirements.txt
    streamlit run app.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import date
from pathlib import Path

# ============================================================
# Configuração
# ============================================================
st.set_page_config(
    page_title="CLEAR 2026 — Organização",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# CSS customizado para garantir tema claro e estilo FGV
st.markdown("""
<style>
    /* Forçar tema claro */
    .stApp { background-color: #FFFFFF; color: #1E293B; }
    [data-testid="stSidebar"] { background-color: #F8FAFC; }
    
    /* Títulos em azul FGV */
    h1, h2, h3 { color: #002B5C !important; }
    
    /* Métricas */
    [data-testid="stMetricValue"] { color: #002B5C; font-weight: 700; }
    [data-testid="stMetricLabel"] { color: #475569; }
    
    /* Abas */
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #F1F5F9;
        border-radius: 6px 6px 0 0;
        padding: 8px 16px;
        color: #475569;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0033A0 !important;
        color: white !important;
    }
    
    /* Botão de filtros */
    .stRadio > label, .stMultiSelect > label, .stSelectbox > label {
        color: #002B5C;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

MASTER_FILE = Path(__file__).parent / "CLEAR_Master_2026.xlsx"

CORES_STATUS = {
    "Concluído": "#1F7A3D",      # verde sóbrio
    "Em Andamento": "#0033A0",   # azul FGV
    "Não Iniciado": "#6B7280",   # cinza médio
    "Atrasado": "#B91C1C",       # vermelho escuro
    "Reunião": "#6B21A8",        # roxo sóbrio
}

# Paleta FGV
COR_PRIMARIA = "#0033A0"     # azul FGV
COR_PRIMARIA_ESCURA = "#002B5C"  # azul FGV escuro
COR_SECUNDARIA = "#475569"   # cinza escuro
COR_FUNDO_CARD = "#F8FAFC"   # cinza muito claro
COR_BORDA = "#E2E8F0"        # cinza claro


@st.cache_data(show_spinner="Carregando planilha mestra...")
def carregar_dados():
    atividades = pd.read_excel(MASTER_FILE, sheet_name="Atividades")
    resp = pd.read_excel(MASTER_FILE, sheet_name="Resp_Atividade")
    alocacao = pd.read_excel(MASTER_FILE, sheet_name="Alocacao")
    envolvimento = pd.read_excel(MASTER_FILE, sheet_name="Envolvimento")
    pessoas = pd.read_excel(MASTER_FILE, sheet_name="Pessoas")
    projetos = pd.read_excel(MASTER_FILE, sheet_name="Projetos")

    atividades["prazo"] = pd.to_datetime(atividades["prazo"], errors="coerce")
    resp["prazo"] = pd.to_datetime(resp["prazo"], errors="coerce")
    alocacao["data_mes"] = pd.to_datetime(alocacao["data_mes"], errors="coerce")
    atividades["eh_entregavel"] = atividades["eh_entregavel"].fillna(False).astype(bool)
    resp["eh_entregavel"] = resp["eh_entregavel"].fillna(False).astype(bool)
    atividades["responsaveis"] = atividades["responsaveis"].fillna("")

    return atividades, resp, alocacao, envolvimento, pessoas, projetos


def aplicar_filtros_globais(df_at, df_resp):
    with st.sidebar:
        st.header("🔍 Filtros")

        st.subheader("Período")
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

        st.subheader("Projetos")
        projetos_disp = sorted(df_at["projeto"].dropna().unique())
        projs_selecionados = st.multiselect(
            "Selecione",
            options=projetos_disp,
            default=projetos_disp,
            label_visibility="collapsed",
        )

        st.subheader("Status")
        status_disp = sorted(df_at["status"].dropna().unique())
        status_selecionados = st.multiselect(
            "Selecione",
            options=status_disp,
            default=[s for s in status_disp if s != "Concluído"],
            label_visibility="collapsed",
        )

        st.markdown("---")
        st.caption(f"📅 Hoje: {date.today().strftime('%d/%m/%Y')}")

    mask_at = df_at["projeto"].isin(projs_selecionados) & df_at["status"].isin(status_selecionados)
    mask_resp = df_resp["projeto"].isin(projs_selecionados) & df_resp["status"].isin(status_selecionados)
    return df_at[mask_at].copy(), df_resp[mask_resp].copy(), data_min, data_max, projs_selecionados


def view_portfolio(df_at, data_min, data_max):
    st.header("📊 Portfólio dos Projetos")
    st.caption(
        "Onde estão os picos de entrega? Quais projetos concentram dedicação?"
    )

    com_data = df_at[df_at["prazo"].notna()].copy()
    no_periodo = com_data[(com_data["prazo"] >= data_min) & (com_data["prazo"] <= data_max)]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Atividades no período", len(no_periodo))
    c2.metric("Entregáveis críticos", int(no_periodo["eh_entregavel"].sum()))
    c3.metric("Em andamento", int((no_periodo["status"] == "Em Andamento").sum()))
    atrasados = no_periodo[
        (no_periodo["status"] != "Concluído")
        & (no_periodo["prazo"] < pd.Timestamp(date.today()))
    ]
    c4.metric("⚠️ Atrasados", len(atrasados), delta_color="inverse")

    st.markdown("---")

    st.subheader("Densidade de atividades — projeto × mês")
    st.caption("Quanto mais escuro, mais atividades com prazo no mês.")
    if not com_data.empty:
        com_data["mes_ref"] = com_data["prazo"].dt.to_period("M").dt.to_timestamp()
        densidade = com_data.groupby(["projeto", "mes_ref"]).size().reset_index(name="qtd")
        fig = px.density_heatmap(
            densidade,
            x="mes_ref",
            y="projeto",
            z="qtd",
            color_continuous_scale=[[0, "#FFFFFF"], [0.3, "#BFDBFE"], [0.6, "#3B82F6"], [1.0, "#002B5C"]],
            labels={"mes_ref": "Mês", "projeto": "", "qtd": "Atividades"},
        )
        fig.update_xaxes(dtick="M1", tickformat="%b/%y")
        fig.update_layout(height=380, margin=dict(t=20, b=20))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Sem atividades com data no filtro atual.")

    st.subheader("Timeline de atividades no período")
    st.caption("Cada ponto é uma atividade. Maior = entregável crítico. Linha vermelha = hoje.")
    if not no_periodo.empty:
        no_periodo = no_periodo.copy()
        no_periodo["tamanho"] = no_periodo["eh_entregavel"].map({True: 18, False: 8})
        no_periodo["atividade_curta"] = no_periodo["atividade"].str[:90]
        fig = px.scatter(
            no_periodo,
            x="prazo",
            y="projeto",
            color="status",
            color_discrete_map=CORES_STATUS,
            size="tamanho",
            size_max=18,
            hover_name="atividade_curta",
            hover_data={
                "responsaveis": True,
                "sub_projeto": True,
                "prazo": "|%d/%m/%Y",
                "tamanho": False,
                "atividade_curta": False,
            },
        )
        fig.update_layout(height=460, margin=dict(t=20, b=20))
        hoje_ts = pd.Timestamp(date.today())
        fig.add_shape(
            type="line", x0=hoje_ts, x1=hoje_ts, y0=0, y1=1, yref="paper",
            line=dict(color="red", width=2, dash="dash"),
        )
        fig.add_annotation(
            x=hoje_ts, y=1, yref="paper", text="hoje", showarrow=False,
            font=dict(color="red"), yshift=10,
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Sem atividades no período filtrado.")

    st.subheader("📌 Próximos entregáveis no período")
    entregaveis = no_periodo[no_periodo["eh_entregavel"]].sort_values("prazo")
    if entregaveis.empty:
        st.info("Sem entregáveis no período filtrado.")
    else:
        st.dataframe(
            entregaveis[
                ["prazo", "projeto", "sub_projeto", "atividade", "responsaveis", "status"]
            ].rename(
                columns={
                    "prazo": "Prazo",
                    "projeto": "Projeto",
                    "sub_projeto": "Subprojeto",
                    "atividade": "Entregável",
                    "responsaveis": "Responsáveis",
                    "status": "Status",
                }
            ),
            use_container_width=True,
            hide_index=True,
            column_config={"Prazo": st.column_config.DateColumn(format="DD/MM/YYYY")},
        )


def view_pesquisador(df_resp, df_alocacao, df_pessoas, data_min, data_max):
    st.header("👤 Visão por Pesquisador")
    st.caption("Pulverização de projetos por mês + atividades sob responsabilidade.")

    pessoas_lista = df_pessoas["pessoa"].sort_values().tolist()
    pessoa = st.selectbox("Pesquisador(a)", pessoas_lista)

    aloc_pessoa = df_alocacao[df_alocacao["pessoa"] == pessoa].sort_values("data_mes").copy()
    if aloc_pessoa.empty:
        st.warning(f"{pessoa} não está na planilha de alocação (mas pode ter atividades atribuídas).")
    else:
        carga_sem = aloc_pessoa["carga_horaria_semanal"].iloc[0]
        media_proj = aloc_pessoa["n_projetos_no_mes"].mean()
        max_proj = aloc_pessoa["n_projetos_no_mes"].max()
        pico_mes = aloc_pessoa.loc[aloc_pessoa["n_projetos_no_mes"].idxmax(), "mes"]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Dedicação contratual", f"{carga_sem:.0f}h/sem")
        c2.metric("Média de projetos/mês", f"{media_proj:.1f}")
        c3.metric("Pico de projetos", int(max_proj), help=f"em {pico_mes}")
        n_at = (df_resp[df_resp["responsavel"] == pessoa]["status"] != "Concluído").sum()
        c4.metric("Atividades em aberto", int(n_at))

        st.subheader("Pulverização: projetos simultâneos por mês")
        st.caption(
            "Quantos projetos a pessoa está envolvida em cada mês. "
            "Acima de 6 projetos simultâneos é sinal de pulverização excessiva."
        )

        def cor_n_projetos(n):
            if n >= 8:
                return "#B91C1C"
            if n >= 6:
                return "#D97706"
            return "#0033A0"

        cores = [cor_n_projetos(n) for n in aloc_pessoa["n_projetos_no_mes"]]
        fig = go.Figure()
        fig.add_trace(
            go.Bar(
                x=aloc_pessoa["mes"],
                y=aloc_pessoa["n_projetos_no_mes"],
                marker_color=cores,
                text=aloc_pessoa["n_projetos_no_mes"],
                textposition="outside",
                hovertemplate="<b>%{x}</b><br>Projetos simultâneos: %{y}<extra></extra>",
            )
        )
        fig.update_layout(
            height=320,
            margin=dict(t=40, b=20),
            yaxis_title="Nº de projetos simultâneos",
            xaxis_title="",
            showlegend=False,
        )
        # Linhas de referência sem add_hline (compatibilidade Plotly 6.7)
        fig.add_shape(type="line", x0=-0.5, x1=len(aloc_pessoa)-0.5, y0=8, y1=8,
                      line=dict(color="red", width=1, dash="dash"))
        fig.add_annotation(x=len(aloc_pessoa)-1, y=8, text="muito pulverizado",
                           showarrow=False, font=dict(color="red", size=10), yshift=10, xanchor="right")
        fig.add_shape(type="line", x0=-0.5, x1=len(aloc_pessoa)-0.5, y0=6, y1=6,
                      line=dict(color="orange", width=1, dash="dot"))
        fig.add_annotation(x=len(aloc_pessoa)-1, y=6, text="atenção",
                           showarrow=False, font=dict(color="orange", size=10), yshift=10, xanchor="right")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    st.subheader(f"Atividades de {pessoa}")
    at_pessoa = df_resp[df_resp["responsavel"] == pessoa].copy()
    at_com_data = at_pessoa[at_pessoa["prazo"].notna()]
    at_no_periodo = at_com_data[
        (at_com_data["prazo"] >= data_min) & (at_com_data["prazo"] <= data_max)
    ]

    if not at_no_periodo.empty:
        at_no_periodo = at_no_periodo.copy()
        at_no_periodo["tamanho"] = at_no_periodo["eh_entregavel"].map({True: 16, False: 8})
        at_no_periodo["atividade_curta"] = at_no_periodo["atividade"].str[:90]
        fig = px.scatter(
            at_no_periodo,
            x="prazo",
            y="projeto",
            color="status",
            color_discrete_map=CORES_STATUS,
            size="tamanho",
            size_max=16,
            hover_name="atividade_curta",
            hover_data={
                "sub_projeto": True,
                "prazo": "|%d/%m/%Y",
                "tamanho": False,
                "atividade_curta": False,
            },
        )
        fig.update_layout(height=360, margin=dict(t=20, b=20))
        hoje_ts = pd.Timestamp(date.today())
        fig.add_shape(
            type="line", x0=hoje_ts, x1=hoje_ts, y0=0, y1=1, yref="paper",
            line=dict(color="red", width=2, dash="dash"),
        )
        fig.add_annotation(
            x=hoje_ts, y=1, yref="paper", text="hoje", showarrow=False,
            font=dict(color="red"), yshift=10,
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info(f"{pessoa} não tem atividades com data no período selecionado.")

    with st.expander(f"Ver tabela completa ({len(at_com_data)} atividades com data)"):
        st.dataframe(
            at_com_data.sort_values("prazo")[
                ["prazo", "projeto", "sub_projeto", "atividade", "status", "eh_entregavel"]
            ].rename(
                columns={
                    "prazo": "Prazo",
                    "projeto": "Projeto",
                    "sub_projeto": "Subprojeto",
                    "atividade": "Atividade",
                    "status": "Status",
                    "eh_entregavel": "Entregável?",
                }
            ),
            use_container_width=True,
            hide_index=True,
            column_config={"Prazo": st.column_config.DateColumn(format="DD/MM/YYYY")},
        )


def view_linha_chegada(df_at, data_min, data_max):
    st.header("🏁 Linha de Chegada")
    st.caption("Calendário de entregas críticas — o que precisa ficar pronto e quando.")

    entregaveis = df_at[df_at["eh_entregavel"] & df_at["prazo"].notna()].copy()
    entregaveis_periodo = entregaveis[
        (entregaveis["prazo"] >= data_min) & (entregaveis["prazo"] <= data_max)
    ].sort_values("prazo")

    if entregaveis_periodo.empty:
        st.info("Nenhum entregável crítico no período. Tente ampliar a janela na barra lateral.")
        return

    hoje = pd.Timestamp(date.today())
    proximos_7 = entregaveis_periodo[
        (entregaveis_periodo["prazo"] >= hoje)
        & (entregaveis_periodo["prazo"] <= hoje + pd.Timedelta(days=7))
    ]
    proximos_30 = entregaveis_periodo[
        (entregaveis_periodo["prazo"] > hoje + pd.Timedelta(days=7))
        & (entregaveis_periodo["prazo"] <= hoje + pd.Timedelta(days=30))
    ]
    atrasados = entregaveis_periodo[
        (entregaveis_periodo["prazo"] < hoje) & (entregaveis_periodo["status"] != "Concluído")
    ]
    futuros = entregaveis_periodo[entregaveis_periodo["prazo"] > hoje + pd.Timedelta(days=30)]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("⚠️ Atrasados", len(atrasados), delta_color="inverse")
    c2.metric("🔥 Próximos 7 dias", len(proximos_7))
    c3.metric("⏳ Próximos 30 dias", len(proximos_30))
    c4.metric("📅 Depois", len(futuros))

    st.markdown("---")

    def mostrar_bloco(titulo, df, emoji, limite=15):
        if df.empty:
            return
        total = len(df)
        st.subheader(f"{emoji} {titulo} ({total})")
        df_mostrar = df.head(limite)
        for _, row in df_mostrar.iterrows():
            dias = (row["prazo"] - hoje).days
            if dias < 0:
                dias_txt = f"<b>{abs(dias)} dias atrasado</b>"
            elif dias == 0:
                dias_txt = "<b>hoje</b>"
            else:
                dias_txt = f"em {dias} dias"
            cor = CORES_STATUS.get(row["status"], "#94A3B8")
            resp_txt = row["responsaveis"] if row["responsaveis"] else "<i>sem responsável CLEAR</i>"
            st.markdown(
                f"""
                <div style="border-left: 4px solid {cor}; padding: 8px 14px; margin-bottom: 8px; background: #F8FAFC; border-radius: 4px;">
                  <div style="font-size: 0.85rem; color: #475569;">
                    <b>{row['prazo'].strftime('%d/%m/%Y')}</b> · {dias_txt} · <span style="color:{cor};">{row['status']}</span>
                  </div>
                  <div style="font-weight: 600; margin: 4px 0;">{row['atividade']}</div>
                  <div style="font-size: 0.85rem; color: #475569;">
                    <i>{row['projeto']}</i> — {row['sub_projeto']}<br>
                    👥 {resp_txt}
                  </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        if total > limite:
            st.caption(f"...e mais {total - limite} item(ns). Use os filtros laterais para refinar.")

    mostrar_bloco("Atrasados", atrasados, "⚠️")
    mostrar_bloco("Próximos 7 dias", proximos_7, "🔥")
    mostrar_bloco("Próximos 30 dias", proximos_30, "⏳")
    mostrar_bloco("Mais à frente", futuros, "📅")


def main():
    st.title("CLEAR 2026 — Organização e Alocação")

    try:
        df_at, df_resp, df_alocacao, df_envolvimento, df_pessoas, df_projetos = carregar_dados()
    except FileNotFoundError:
        st.error(
            f"Planilha mestra não encontrada em {MASTER_FILE}. "
            "Coloque CLEAR_Master_2026.xlsx na mesma pasta do app.py."
        )
        st.stop()

    df_at_f, df_resp_f, data_min, data_max, projs = aplicar_filtros_globais(df_at, df_resp)

    tab1, tab2, tab3 = st.tabs(["📊 Portfólio", "👤 Pesquisador", "🏁 Linha de Chegada"])
    with tab1:
        view_portfolio(df_at_f, data_min, data_max)
    with tab2:
        view_pesquisador(df_resp_f, df_alocacao, df_pessoas, data_min, data_max)
    with tab3:
        view_linha_chegada(df_at_f, data_min, data_max)

    st.markdown("---")
    st.caption(
        "Dados: `CLEAR_Master_2026.xlsx`. "
        "Para atualizar, edite a planilha e dê reload no app (botão R)."
    )


if __name__ == "__main__":
    main()
