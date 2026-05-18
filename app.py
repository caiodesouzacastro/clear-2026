"""
CLEAR 2026 — Dashboard de Organização e Alocação
FGV CLEAR

4 abas: Visão Geral (Gantt), Calendário, Equipe, Pesquisador.
Gantt e calendário em HTML puro (sem Plotly) para carregar rápido.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date
from pathlib import Path
import calendar as cal_mod

# ============================================================
# Configuração
# ============================================================
st.set_page_config(page_title="CLEAR 2026", page_icon="📊", layout="wide",
                   initial_sidebar_state="expanded")

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
AZUL_ESCURO = "#26456B"

CORES_STATUS = {
    "Concluído": "#6F7E8C",
    "Em Andamento": "#5090D3",
    "Não Iniciado": "#3A5169",
    "Atrasado": "#7BB3F0",
    "Reunião": "#4A6FA5",
}

MESES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
            "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MESES_PT_FULL = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                 "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

# ============================================================
# CSS — tema escuro Bloomberg
# ============================================================
st.markdown(f"""
<style>
    .stApp {{ background-color: {BG}; }}
    .main, .block-container {{ background-color: {BG}; color: {TEXTO}; }}
    .block-container {{ padding-top: 2rem; }}
    [data-testid="stSidebar"] {{ background-color: {BG_CARD}; }}
    [data-testid="stSidebar"] > div {{ background-color: {BG_CARD}; }}
    h1, h2, h3, h4, h5, h6 {{ color: {TEXTO} !important; font-weight: 600; }}
    p, span, label {{ color: {TEXTO}; }}
    [data-testid="stMetricValue"] {{ color: {TEXTO} !important; font-weight: 700; font-size: 2rem; }}
    [data-testid="stMetricLabel"] {{ color: {TEXTO_DIM} !important; font-size: 0.8rem;
        text-transform: uppercase; letter-spacing: 0.05em; }}
    .stTabs [data-baseweb="tab-list"] {{ gap: 4px; background-color: {BG};
        border-bottom: 1px solid {BORDA}; }}
    .stTabs [data-baseweb="tab"] {{ background-color: transparent; color: {TEXTO_DIM};
        padding: 10px 20px; border-bottom: 2px solid transparent; }}
    .stTabs [aria-selected="true"] {{ background-color: transparent !important;
        color: {AZUL} !important; border-bottom: 2px solid {AZUL} !important; font-weight: 600; }}
    [data-baseweb="tag"] {{ background-color: {BG_HOVER} !important; color: {TEXTO} !important; }}
    [data-baseweb="select"] > div {{ background-color: {BG_CARD} !important;
        border-color: {BORDA} !important; }}
    [data-baseweb="input"] > div {{ background-color: {BG_CARD} !important; }}
    [data-testid="stDataFrame"] {{ background-color: {BG_CARD}; }}
    hr {{ border-color: {BORDA}; margin: 1rem 0; }}
    header[data-testid="stHeader"] {{ background-color: {BG}; }}
    .stButton button {{ background-color: {BG_CARD}; color: {TEXTO};
        border: 1px solid {BORDA}; }}
    .stButton button:hover {{ background-color: {BG_HOVER}; border-color: {AZUL}; }}

    /* Forçar azul nos controles nativos (remove o laranja padrão do Streamlit) */
    [data-testid="stRadio"] [data-baseweb="radio"] div[aria-checked="true"],
    [data-baseweb="radio"] div[aria-checked="true"] {{
        background-color: {AZUL} !important;
        border-color: {AZUL} !important;
    }}
    [data-testid="stCheckbox"] [data-baseweb="checkbox"] span[aria-checked="true"] {{
        background-color: {AZUL} !important;
        border-color: {AZUL} !important;
    }}
    input[type="radio"]:checked + div {{ background-color: {AZUL} !important; }}
    [data-testid="stSlider"] [data-baseweb="slider"] div {{ background-color: {AZUL} !important; }}
    a {{ color: {AZUL_CLARO} !important; }}
    *:focus {{ outline-color: {AZUL} !important; }}
    [data-baseweb="radio"] svg {{ fill: {AZUL} !important; }}
    [role="radiogroup"] label div:first-child {{ border-color: {AZUL} !important; }}
    /* Rede de segurança: reescreve a cor primária via variável CSS do Streamlit */
    :root {{
        --primary-color: {AZUL} !important;
        --primaryColor: {AZUL} !important;
    }}
</style>
""", unsafe_allow_html=True)


# ============================================================
# Dados
# ============================================================
@st.cache_data(show_spinner="Carregando...")
def carregar_dados():
    at = pd.read_excel(MASTER_FILE, sheet_name="Atividades")
    resp = pd.read_excel(MASTER_FILE, sheet_name="Resp_Atividade")
    aloc = pd.read_excel(MASTER_FILE, sheet_name="Alocacao")
    pes = pd.read_excel(MASTER_FILE, sheet_name="Pessoas")
    at["prazo"] = pd.to_datetime(at["prazo"], errors="coerce")
    resp["prazo"] = pd.to_datetime(resp["prazo"], errors="coerce")
    aloc["data_mes"] = pd.to_datetime(aloc["data_mes"], errors="coerce")
    at["eh_entregavel"] = at["eh_entregavel"].fillna(False).astype(bool)
    resp["eh_entregavel"] = resp["eh_entregavel"].fillna(False).astype(bool)
    at["responsaveis"] = at["responsaveis"].fillna("")
    return at, resp, aloc, pes


# ============================================================
# Sidebar — filtros globais
# ============================================================
def sidebar_filtros(df_at):
    with st.sidebar:
        st.markdown("### CLEAR 2026")
        st.caption("Organização e Alocação")
        st.markdown("---")

        st.markdown("**Mostrar**")
        escopo = st.radio(
            "escopo",
            ["Só entregáveis críticos", "Todas as atividades"],
            label_visibility="collapsed",
        )

        st.markdown("**Projetos**")
        projetos_disp = sorted(df_at["projeto"].dropna().unique())
        projs = st.multiselect("projs", options=projetos_disp, default=projetos_disp,
                               label_visibility="collapsed")

        st.markdown("---")
        st.caption(f"Hoje: {date.today().strftime('%d/%m/%Y')}")
        st.caption("Dados: CLEAR_Master_2026.xlsx")

    return escopo, projs


def filtrar(df_at, escopo, projs):
    d = df_at[df_at["projeto"].isin(projs)].copy()
    if escopo == "Só entregáveis críticos":
        d = d[d["eh_entregavel"]]
    return d


# ============================================================
# ABA 1 — Visão Geral (Gantt em HTML)
# ============================================================
def aba_gantt(df_at):
    st.subheader("Visão Geral do Ano")

    com_data = df_at[df_at["prazo"].notna()].copy()
    hoje = pd.Timestamp(date.today())

    # KPIs
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Atividades", len(com_data))
    c2.metric("Entregáveis", int(com_data["eh_entregavel"].sum()))
    c3.metric("Em andamento", int((com_data["status"] == "Em Andamento").sum()))
    atrasadas = com_data[(com_data["status"] != "Concluído") & (com_data["prazo"] < hoje)]
    c4.metric("Atrasadas", len(atrasadas))

    st.markdown("---")
    st.markdown("**Linha do tempo por projeto** — 2026")

    if com_data.empty:
        st.info("Sem atividades com data no filtro atual.")
        return

    # Agrupar por projeto: data mínima e máxima de cada
    ano_ini = pd.Timestamp("2026-01-01")
    ano_fim = pd.Timestamp("2026-12-31")
    total_dias = (ano_fim - ano_ini).days

    projetos = sorted(com_data["projeto"].unique())

    # Cabeçalho de meses
    header_meses = "".join(
        f'<div style="flex:1; text-align:center; font-size:0.7rem; '
        f'color:{TEXTO_DIM2}; border-left:1px solid {BORDA}; padding:2px 0;">{m}</div>'
        for m in MESES_PT
    )

    linhas_html = ""
    for proj in projetos:
        d = com_data[com_data["projeto"] == proj]
        dmin = d["prazo"].min()
        dmax = d["prazo"].max()
        ini_pct = max(0, (dmin - ano_ini).days / total_dias * 100)
        fim_pct = min(100, (dmax - ano_ini).days / total_dias * 100)
        larg_pct = max(1.5, fim_pct - ini_pct)
        n_entreg = int(d["eh_entregavel"].sum())
        n_total = len(d)

        marcadores = ""
        for _, r in d[d["eh_entregavel"]].iterrows():
            pos = (r["prazo"] - ano_ini).days / total_dias * 100
            tt = str(r["atividade"])[:60].replace('"', "'")
            marcadores += (f'<div title="{tt} ({r["prazo"].strftime("%d/%m")})" style="position:absolute;left:{pos:.1f}%;top:50%;transform:translate(-50%,-50%);width:9px;height:9px;background:{AZUL_CLARO};border:1.5px solid {BG};border-radius:50%;z-index:2;"></div>')

        linhas_html += (
            f'<div style="display:flex;align-items:center;margin-bottom:6px;">'
            f'<div style="width:170px;font-size:0.8rem;color:{TEXTO};padding-right:10px;text-align:right;flex-shrink:0;">'
            f'{proj}<br><span style="color:{TEXTO_DIM2};font-size:0.7rem;">{n_total} ativ · {n_entreg} entreg</span></div>'
            f'<div style="flex:1;position:relative;height:26px;background:{BG_CARD};border-radius:3px;">'
            f'<div style="position:absolute;left:{ini_pct:.1f}%;width:{larg_pct:.1f}%;top:6px;height:14px;background:{AZUL_ESCURO};border-radius:3px;"></div>'
            f'{marcadores}</div></div>'
        )

    html = (
        f'<div style="background:{BG};padding:14px;border-radius:6px;border:1px solid {BORDA};">'
        f'<div style="display:flex;margin-bottom:8px;">'
        f'<div style="width:170px;flex-shrink:0;"></div>'
        f'<div style="flex:1;display:flex;">{header_meses}</div></div>'
        f'{linhas_html}</div>'
        f'<div style="font-size:0.75rem;color:{TEXTO_DIM2};margin-top:8px;">'
        f'● Marcadores azuis = entregáveis críticos · Barra = período de atividades do projeto</div>'
    )
    st.markdown(html, unsafe_allow_html=True)


# ============================================================
# ABA 2 — Calendário mensal (HTML)
# ============================================================
def aba_calendario(df_at):
    st.subheader("Calendário de Entregas")

    com_data = df_at[df_at["prazo"].notna()].copy()
    if com_data.empty:
        st.info("Sem atividades com data no filtro atual.")
        return

    # Navegação de mês
    if "cal_mes" not in st.session_state:
        hoje = date.today()
        st.session_state.cal_mes = hoje.month if hoje.year == 2026 else 1

    c1, c2, c3 = st.columns([1, 3, 1])
    with c1:
        if st.button("← Anterior", use_container_width=True):
            st.session_state.cal_mes = max(1, st.session_state.cal_mes - 1)
    with c3:
        if st.button("Próximo →", use_container_width=True):
            st.session_state.cal_mes = min(12, st.session_state.cal_mes + 1)
    with c2:
        st.markdown(
            f"<h3 style='text-align:center; margin:0;'>"
            f"{MESES_PT_FULL[st.session_state.cal_mes - 1]} 2026</h3>",
            unsafe_allow_html=True,
        )

    mes = st.session_state.cal_mes
    do_mes = com_data[(com_data["prazo"].dt.year == 2026) & (com_data["prazo"].dt.month == mes)]

    # Montar grade do calendário
    primeiro_dia, num_dias = cal_mod.monthrange(2026, mes)
    # primeiro_dia: 0=segunda ... 6=domingo
    dias_semana = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
    header = "".join(
        f'<div style="flex:1; text-align:center; font-size:0.75rem; '
        f'color:{TEXTO_DIM2}; padding:4px 0; font-weight:600;">{d}</div>'
        for d in dias_semana
    )

    hoje = date.today()
    celulas = []
    # espaços vazios antes do dia 1
    for _ in range(primeiro_dia):
        celulas.append('<div style="flex:1; min-height:90px;"></div>')

    for dia in range(1, num_dias + 1):
        entregas_dia = do_mes[do_mes["prazo"].dt.day == dia]
        eh_hoje = (hoje.year == 2026 and hoje.month == mes and hoje.day == dia)
        borda = f"2px solid {AZUL_CLARO}" if eh_hoje else f"1px solid {BORDA}"

        itens = ""
        for _, r in entregas_dia.head(4).iterrows():
            cor = CORES_STATUS.get(r["status"], TEXTO_DIM)
            marca = "◆ " if r["eh_entregavel"] else ""
            itens += (
                f'<div title="{r["atividade"][:80]}" style="font-size:0.65rem; '
                f'background:{BG_HOVER}; border-left:2px solid {cor}; '
                f'padding:1px 3px; margin-bottom:2px; border-radius:2px; '
                f'white-space:nowrap; overflow:hidden; text-overflow:ellipsis; '
                f'color:{TEXTO};">{marca}{r["atividade"][:22]}</div>'
            )
        if len(entregas_dia) > 4:
            itens += (f'<div style="font-size:0.62rem; color:{TEXTO_DIM2};">'
                      f'+{len(entregas_dia) - 4} mais</div>')

        celulas.append(
            f'<div style="flex:1; min-height:90px; border:{borda}; '
            f'border-radius:3px; padding:3px; background:{BG_CARD}; '
            f'margin:1px; overflow:hidden;">'
            f'<div style="font-size:0.75rem; color:{TEXTO_DIM}; '
            f'font-weight:600; margin-bottom:2px;">{dia}</div>{itens}</div>'
        )

    # quebrar em semanas (linhas de 7)
    linhas = ""
    for i in range(0, len(celulas), 7):
        semana = celulas[i:i + 7]
        while len(semana) < 7:
            semana.append('<div style="flex:1;"></div>')
        linhas += f'<div style="display:flex;">{"".join(semana)}</div>'

    html = (
        f'<div style="background:{BG};padding:10px;border-radius:6px;border:1px solid {BORDA};">'
        f'<div style="display:flex;margin-bottom:4px;">{header}</div>'
        f'{linhas}</div>'
        f'<div style="font-size:0.75rem;color:{TEXTO_DIM2};margin-top:8px;">'
        f'◆ = entregável crítico · Cor da borda do item = status</div>'
    )
    st.markdown(html, unsafe_allow_html=True)

    st.metric("Entregas neste mês", len(do_mes))


# ============================================================
# ABA 3 — Equipe (mapa de calor pessoa × mês)
# ============================================================
def aba_equipe(df_aloc):
    st.subheader("Equipe — Carga por Mês")
    st.caption("Número de projetos simultâneos. Quanto mais claro, mais pulverizado.")

    if df_aloc.empty:
        st.info("Sem dados de alocação.")
        return

    pessoas = sorted(df_aloc["pessoa"].unique())
    meses_ordem = ["Maio", "Junho", "Julho", "Agosto", "Setembro",
                   "Outubro", "Novembro", "Dezembro"]

    # cor da célula conforme nº de projetos
    def cor_celula(n):
        if n == 0:
            return BG_CARD
        if n <= 3:
            return "#26456B"
        if n <= 5:
            return "#3D6A9E"
        if n <= 7:
            return "#5090D3"
        return "#7BB3F0"

    # cabeçalho
    header = f'<div style="width:90px; flex-shrink:0;"></div>'
    for m in meses_ordem:
        header += (f'<div style="flex:1; text-align:center; font-size:0.7rem; '
                   f'color:{TEXTO_DIM2}; padding:4px 0;">{m[:3]}</div>')

    linhas = ""
    for p in pessoas:
        dp = df_aloc[df_aloc["pessoa"] == p]
        celulas = (f'<div style="width:90px; flex-shrink:0; font-size:0.8rem; '
                   f'color:{TEXTO}; padding-right:8px; text-align:right;">{p}</div>')
        for m in meses_ordem:
            linha_m = dp[dp["mes"] == m]
            n = int(linha_m["n_projetos_no_mes"].iloc[0]) if not linha_m.empty else 0
            cor = cor_celula(n)
            txt_cor = BG if n > 5 else TEXTO_DIM
            celulas += (f'<div style="flex:1; text-align:center; padding:8px 0; '
                        f'margin:1px; background:{cor}; border-radius:3px; '
                        f'font-size:0.8rem; color:{txt_cor}; font-weight:600;">'
                        f'{n if n > 0 else ""}</div>')
        linhas += f'<div style="display:flex; align-items:center; margin-bottom:2px;">{celulas}</div>'

    html = (
        f'<div style="background:{BG};padding:12px;border-radius:6px;border:1px solid {BORDA};">'
        f'<div style="display:flex;margin-bottom:4px;">{header}</div>'
        f'{linhas}</div>'
        f'<div style="font-size:0.75rem;color:{TEXTO_DIM2};margin-top:8px;">'
        f'Número = projetos simultâneos no mês · 6+ projetos indica pulverização</div>'
    )
    st.markdown(html, unsafe_allow_html=True)


# ============================================================
# ABA 4 — Pesquisador individual
# ============================================================
def aba_pesquisador(df_resp, df_aloc, df_pessoas):
    st.subheader("Visão por Pesquisador")

    pessoas_lista = df_pessoas["pessoa"].sort_values().tolist()
    pessoa = st.selectbox("Pesquisador(a)", pessoas_lista)

    aloc_p = df_aloc[df_aloc["pessoa"] == pessoa].sort_values("data_mes")
    if not aloc_p.empty:
        carga = aloc_p["carga_horaria_semanal"].iloc[0]
        media = aloc_p["n_projetos_no_mes"].mean()
        pico = aloc_p["n_projetos_no_mes"].max()
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Dedicação", f"{carga:.0f}h/sem")
        c2.metric("Média projetos/mês", f"{media:.1f}")
        c3.metric("Pico de projetos", int(pico))
        n_aberto = (df_resp[df_resp["responsavel"] == pessoa]["status"] != "Concluído").sum()
        c4.metric("Atividades em aberto", int(n_aberto))

        st.markdown("---")
        st.markdown("**Projetos simultâneos por mês**")

        def cor_n(n):
            if n >= 8:
                return AZUL_CLARO
            if n >= 6:
                return AZUL
            return AZUL_ESCURO

        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=aloc_p["mes"], y=aloc_p["n_projetos_no_mes"],
            marker_color=[cor_n(n) for n in aloc_p["n_projetos_no_mes"]],
            text=aloc_p["n_projetos_no_mes"], textposition="outside",
            textfont=dict(color=TEXTO),
            hovertemplate="<b>%{x}</b><br>%{y} projetos<extra></extra>",
        ))
        fig.update_layout(
            height=300, margin=dict(t=20, b=20, l=10, r=10),
            paper_bgcolor=BG, plot_bgcolor=BG,
            font=dict(color=TEXTO_DIM), showlegend=False,
            xaxis=dict(gridcolor=BORDA), yaxis=dict(gridcolor=BORDA),
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    else:
        st.info(f"{pessoa} não está na planilha de alocação.")

    st.markdown("---")
    st.markdown(f"**Atividades de {pessoa}**")
    at_p = df_resp[df_resp["responsavel"] == pessoa].copy()
    at_p = at_p[at_p["prazo"].notna()].sort_values("prazo")
    if at_p.empty:
        st.info(f"{pessoa} não tem atividades com data atribuídas.")
    else:
        st.dataframe(
            at_p[["prazo", "projeto", "sub_projeto", "atividade", "status", "eh_entregavel"]].rename(
                columns={"prazo": "Prazo", "projeto": "Projeto", "sub_projeto": "Subprojeto",
                         "atividade": "Atividade", "status": "Status", "eh_entregavel": "Entregável?"}),
            use_container_width=True, hide_index=True,
            column_config={"Prazo": st.column_config.DateColumn(format="DD/MM/YYYY")},
        )


# ============================================================
# Main
# ============================================================
def main():
    try:
        df_at, df_resp, df_aloc, df_pessoas = carregar_dados()
    except FileNotFoundError:
        st.error("Planilha mestra não encontrada.")
        st.stop()

    escopo, projs = sidebar_filtros(df_at)
    df_at_f = filtrar(df_at, escopo, projs)
    df_resp_f = df_resp[df_resp["projeto"].isin(projs)].copy()
    if escopo == "Só entregáveis críticos":
        df_resp_f = df_resp_f[df_resp_f["eh_entregavel"]]

    tab1, tab2, tab3, tab4 = st.tabs(
        ["Visão Geral", "Calendário", "Equipe", "Pesquisador"])
    with tab1:
        aba_gantt(df_at_f)
    with tab2:
        aba_calendario(df_at_f)
    with tab3:
        aba_equipe(df_aloc)
    with tab4:
        aba_pesquisador(df_resp_f, df_aloc, df_pessoas)


if __name__ == "__main__":
    main()
