"""
CLEAR 2026 — Dashboard de Organização e Alocação
FGV CLEAR

4 abas: Visão Geral (Gantt), Calendário, Equipe, Pesquisador.
Gantt e calendário em HTML puro (sem Plotly) para carregar rápido.
"""

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import date
from pathlib import Path
import calendar as cal_mod

# ============================================================
# Configuração
# ============================================================
st.set_page_config(page_title="CLEAR 2026", page_icon="📊", layout="wide",
                   initial_sidebar_state="expanded")


# ============================================================
# Autenticação (senha simples via st.secrets)
# ============================================================
def _check_password():
    """Tela de senha. Senha vai em .streamlit/secrets.toml como app_password."""
    if st.session_state.get("auth_ok"):
        return True

    # Lê a senha do secrets (Streamlit Cloud) ou variável de ambiente local
    try:
        senha_correta = st.secrets["app_password"]
    except Exception:
        import os
        senha_correta = os.environ.get("APP_PASSWORD", "")

    # Centraliza a tela de login
    _, c, _ = st.columns([1, 1, 1])
    with c:
        st.markdown(
            "<div style='margin-top:80px;text-align:center;'>"
            "<h2 style='color:#5090D3;margin-bottom:4px;'>CLEAR 2026</h2>"
            "<p style='color:#B2BAC2;font-size:14px;margin-top:0;'>"
            "Painel de portfólio · acesso restrito</p></div>",
            unsafe_allow_html=True,
        )
        senha = st.text_input("Senha", type="password",
                              label_visibility="collapsed",
                              placeholder="Digite a senha")
        if senha:
            if senha == senha_correta:
                st.session_state["auth_ok"] = True
                st.rerun()
            else:
                st.error("Senha incorreta.")
    st.stop()


_check_password()

MASTER_FILE = Path(__file__).parent / "CLEAR_Master_2026.xlsx"
FERIAS_FILE = Path(__file__).parent / "Ferias_CLEAR_2026.xlsx"

# Equipe canônica (16, com Fred e Hisrael pelos nomes reais)
EQUIPE = ["Bia B", "Bia S", "Caio", "Carol", "Cecilia", "Fabrícia", "Fred",
          "Hisrael", "Julia", "Junior", "Lorena", "Luan", "Luigi", "Lycia",
          "Michel", "Samu"]

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

# Cor de destaque para entregas de Comunicação no calendário
COR_COMUNICACAO = "#E5A84B"

MESES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
            "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MESES_PT_FULL = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                 "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

# Paleta categórica (legível sobre o navy) — usada nos Gantts de alocação
PALETA = ["#5090D3", "#5DCAA5", "#E5A84B", "#C77DD6", "#E57373", "#7FB8E6",
          "#9CCC65", "#F0A868", "#64B6AC", "#B39DDB", "#4FC3F7", "#FF8A65",
          "#A1C181", "#E8A0BF", "#76C7C0", "#D4A55E"]


def _esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _trunc(s, n):
    s = str(s)
    return s if len(s) <= n else s[:n - 1] + "…"


def render_gantt(groups, color_map, m_min, m_max, eixo_label):
    """Gera um Gantt em SVG. groups: lista de (rótulo, [(sub, de, ate)])."""
    PAD, HDR, NAME_W, COL_W = 12, 40, 180, 64
    BAR_H, BAR_GAP, ROW_PAD, MIN_ROW = 13, 3, 8, 32
    months = list(range(m_min, m_max + 1))
    nM = len(months)
    row_h = [max(MIN_ROW, len(it) * (BAR_H + BAR_GAP) + ROW_PAD) for _, it in groups]
    total_h = HDR + sum(row_h) + PAD
    total_w = NAME_W + nM * COL_W + PAD * 2

    def xm(m):
        return PAD + NAME_W + (m - m_min) * COL_W

    s = [f'<svg viewBox="0 0 {total_w} {total_h}" width="{total_w}" height="{total_h}" '
         f'xmlns="http://www.w3.org/2000/svg" '
         f'style="max-width:100%;height:auto;display:block;margin:0 auto;'
         f'font-family:-apple-system,Segoe UI,sans-serif;">']
    s.append(f'<rect width="{total_w}" height="{total_h}" fill="{BG_CARD}" rx="6"/>')
    for i, m in enumerate(months):
        x = xm(m)
        if i % 2 == 0:
            s.append(f'<rect x="{x}" y="{HDR}" width="{COL_W}" height="{total_h - HDR - PAD}" '
                     f'fill="rgba(255,255,255,0.025)"/>')
        s.append(f'<text x="{x + COL_W / 2}" y="24" text-anchor="middle" font-size="11" '
                 f'fill="{TEXTO_DIM2}">{MESES_PT[m - 1]}</text>')
        s.append(f'<line x1="{x}" y1="30" x2="{x}" y2="{total_h - PAD}" '
                 f'stroke="{BORDA}" stroke-width="0.5"/>')
    xend = PAD + NAME_W + nM * COL_W
    s.append(f'<line x1="{xend}" y1="30" x2="{xend}" y2="{total_h - PAD}" stroke="{BORDA}" stroke-width="0.5"/>')
    s.append(f'<line x1="{PAD + NAME_W}" y1="{HDR - 8}" x2="{PAD + NAME_W}" y2="{total_h - PAD}" '
             f'stroke="{BORDA}" stroke-width="1"/>')
    s.append(f'<text x="{PAD + 4}" y="24" font-size="11" fill="{TEXTO_DIM2}">{eixo_label}</text>')

    y = HDR
    for gi, ((label, items), rh) in enumerate(zip(groups, row_h)):
        # faixa de fundo alternada (separa visualmente cada pessoa/projeto)
        if gi % 2 == 1:
            s.append(f'<rect x="{PAD}" y="{y}" width="{total_w - 2 * PAD}" height="{rh}" '
                     f'fill="rgba(255,255,255,0.035)"/>')
        # divisor superior do grupo (mais forte)
        s.append(f'<line x1="{PAD}" y1="{y}" x2="{total_w - PAD}" y2="{y}" '
                 f'stroke="{AZUL_ESCURO}" stroke-width="1"/>')
        s.append(f'<text x="{PAD + 6}" y="{y + rh / 2 + 4}" font-size="12.5" fill="{TEXTO}" '
                 f'font-weight="500">{_esc(_trunc(label, 26))}<title>{_esc(label)}</title></text>')
        by = y + ROW_PAD / 2
        for (sub, de, ate) in items:
            cor = color_map.get(sub, AZUL)
            bx = xm(de)
            bw = (ate - de + 1) * COL_W - 4
            s.append(f'<g><rect x="{bx + 2}" y="{by}" width="{max(bw, 8)}" height="{BAR_H}" '
                     f'rx="3" fill="{cor}" opacity="0.92"/>'
                     f'<title>{_esc(sub)} · {MESES_PT[de - 1]}–{MESES_PT[ate - 1]}</title>')
            if bw >= 46:
                s.append(f'<text x="{bx + 8}" y="{by + BAR_H - 3}" font-size="10.5" fill="{BG}" '
                         f'font-weight="600">{_esc(_trunc(sub, int(bw / 7)))}</text>')
            s.append('</g>')
            by += BAR_H + BAR_GAP
        y += rh
    # divisor de fechamento
    s.append(f'<line x1="{PAD}" y1="{y}" x2="{total_w - PAD}" y2="{y}" '
             f'stroke="{AZUL_ESCURO}" stroke-width="1"/>')
    s.append('</svg>')
    return "".join(s), total_h

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
    gantt = pd.read_excel(MASTER_FILE, sheet_name="Alocacao_Gantt")
    gantt["mes_de"] = gantt["mes_de"].astype(int)
    gantt["mes_ate"] = gantt["mes_ate"].astype(int)
    at["prazo"] = pd.to_datetime(at["prazo"], errors="coerce")
    resp["prazo"] = pd.to_datetime(resp["prazo"], errors="coerce")
    aloc["data_mes"] = pd.to_datetime(aloc["data_mes"], errors="coerce")
    at["eh_entregavel"] = at["eh_entregavel"].fillna(False).astype(bool)
    resp["eh_entregavel"] = resp["eh_entregavel"].fillna(False).astype(bool)
    at["responsaveis"] = at["responsaveis"].fillna("")
    return at, resp, aloc, pes, gantt


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
    com_data = com_data[(com_data["prazo"] >= pd.Timestamp("2026-01-01")) &
                        (com_data["prazo"] <= pd.Timestamp("2026-12-31"))]
    hoje = pd.Timestamp(date.today())

    # KPIs
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Atividades", len(com_data))
    c2.metric("Entregáveis", int(com_data["eh_entregavel"].sum()))
    c3.metric("Em andamento", int((com_data["status"] == "Em Andamento").sum()))
    atrasadas = com_data[(com_data["status"] != "Concluído") & (com_data["prazo"] < hoje)]
    c4.metric("Atrasadas", len(atrasadas))

    st.markdown("---")

    # Filtro de projetos inline
    todos_projetos = sorted(com_data["projeto"].unique())
    filtro_projs = st.multiselect(
        "Filtrar projetos na timeline",
        options=todos_projetos,
        default=todos_projetos,
        placeholder="Todos os projetos",
    )
    if filtro_projs:
        com_data = com_data[com_data["projeto"].isin(filtro_projs)]

    st.markdown("**Linha do tempo por projeto** — 2026")

    if com_data.empty:
        st.info("Sem atividades com data no filtro atual.")
        return

    ano_ini = pd.Timestamp("2026-01-01")
    ano_fim = pd.Timestamp("2026-12-31")
    total_dias = (ano_fim - ano_ini).days

    projetos = sorted(com_data["projeto"].unique())

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

        # % concluído
        n_conc = int((d["status"] == "Concluído").sum())
        pct_conc = int(n_conc / n_total * 100) if n_total > 0 else 0
        # cor da barra de progresso: verde se >=80, âmbar se >=40, vermelho se <40
        cor_prog = ("#5a9367" if pct_conc >= 80 else "#d99a2b" if pct_conc >= 40 else "#d9534f")

        marcadores = ""
        for _, r in d[d["eh_entregavel"]].iterrows():
            pos = (r["prazo"] - ano_ini).days / total_dias * 100
            tt = str(r["atividade"])[:60].replace('"', "'")
            marcadores += (
                f'<div title="{tt} ({r["prazo"].strftime("%d/%m")})" '
                f'style="position:absolute;left:{pos:.1f}%;top:50%;transform:translate(-50%,-50%);'
                f'width:9px;height:9px;background:{AZUL_CLARO};border:1.5px solid {BG};'
                f'border-radius:50%;z-index:2;"></div>'
            )

        linhas_html += (
            f'<div style="display:flex;align-items:center;margin-bottom:6px;">'
            f'<div style="width:180px;font-size:0.8rem;color:{TEXTO};padding-right:10px;text-align:right;flex-shrink:0;">'
            f'{proj}<br>'
            f'<span style="color:{TEXTO_DIM2};font-size:0.7rem;">{n_total} ativ · {n_entreg} entreg</span><br>'
            f'<div style="margin-top:2px;background:{BORDA};border-radius:3px;height:4px;overflow:hidden;">'
            f'<div style="width:{pct_conc}%;background:{cor_prog};height:4px;border-radius:3px;"></div></div>'
            f'<span style="font-size:0.65rem;color:{cor_prog};">{pct_conc}% concluído</span>'
            f'</div>'
            f'<div style="flex:1;position:relative;height:26px;background:{BG_CARD};border-radius:3px;">'
            f'<div style="position:absolute;left:{ini_pct:.1f}%;width:{larg_pct:.1f}%;top:6px;height:14px;'
            f'background:{AZUL_ESCURO};border-radius:3px;"></div>'
            f'{marcadores}</div></div>'
        )

    html = (
        f'<div style="background:{BG};padding:14px;border-radius:6px;border:1px solid {BORDA};">'
        f'<div style="display:flex;margin-bottom:8px;">'
        f'<div style="width:180px;flex-shrink:0;"></div>'
        f'<div style="flex:1;display:flex;">{header_meses}</div></div>'
        f'{linhas_html}</div>'
        f'<div style="font-size:0.75rem;color:{TEXTO_DIM2};margin-top:8px;">'
        f'● Marcadores azuis = entregáveis críticos · Barra = período · Barra fina = % concluído</div>'
    )
    st.markdown(html, unsafe_allow_html=True)


# ============================================================
# ABA 2 — Calendário mensal (HTML)
# ============================================================
def aba_calendario(df_at):
    st.subheader("Calendário de Entregas")

    com_data = df_at[df_at["prazo"].notna()].copy()
    com_data = com_data[(com_data["prazo"] >= pd.Timestamp("2026-01-01")) &
                        (com_data["prazo"] <= pd.Timestamp("2026-12-31"))]
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
            eh_com = r["projeto"] == "Comunicação"
            cor = COR_COMUNICACAO if eh_com else CORES_STATUS.get(r["status"], TEXTO_DIM)
            largura_borda = "3px" if eh_com else "2px"
            marca = "◆ " if r["eh_entregavel"] else ""
            itens += (
                f'<div title="{r["atividade"][:80]}" style="font-size:0.65rem; '
                f'background:{BG_HOVER}; border-left:{largura_borda} solid {cor}; '
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
        f'◆ = entregável crítico · Cor da borda do item = status · '
        f'<span style="display:inline-block;width:9px;height:9px;'
        f'background:{COR_COMUNICACAO};border-radius:2px;margin:0 3px -1px 0;"></span>'
        f'Comunicação</div>'
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
        n_aberto = (df_resp[df_resp["pessoa"] == pessoa]["status"] != "Concluído").sum()
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
    # df_resp tem id_atividade, pessoa, projeto, prazo, status, eh_entregavel.
    # Para sub_projeto e atividade precisamos juntar com df_at via id_atividade.
    at_p = df_resp[df_resp["pessoa"] == pessoa].copy()
    at_p = at_p[at_p["prazo"].notna()].sort_values("prazo")
    if at_p.empty:
        st.info(f"{pessoa} não tem atividades com data atribuídas.")
    else:
        # Carrega Atividades pra ter sub_projeto e nome da atividade
        df_at_full = pd.read_excel(MASTER_FILE, sheet_name="Atividades",
                                    usecols=["id_atividade", "sub_projeto", "atividade"])
        at_p = at_p.merge(df_at_full, on="id_atividade", how="left")
        st.dataframe(
            at_p[["prazo", "projeto", "sub_projeto", "atividade", "status", "eh_entregavel"]].rename(
                columns={"prazo": "Prazo", "projeto": "Projeto", "sub_projeto": "Subprojeto",
                         "atividade": "Atividade", "status": "Status", "eh_entregavel": "Entregável?"}),
            use_container_width=True, hide_index=True,
            column_config={"Prazo": st.column_config.DateColumn(format="DD/MM/YYYY")},
        )


# ============================================================
# ABA 4 — Farol de Intensidade de Pessoal
# ============================================================
def aba_farol(df_aloc):
    st.subheader("Farol de Intensidade — Carga por Pessoa")
    st.caption("Comparação entre carga **esperada** (definida pelos líderes nos cronogramas) "
               "e carga **percebida** (auto-declarada pela equipe).")

    if df_aloc.empty:
        st.info("Sem dados de alocação.")
        return

    # ── Carrega Carga ESPERADA (do master) ───────────────────────────
    try:
        carga_esp = pd.read_excel(MASTER_FILE, sheet_name="Carga_Esperada")
    except Exception:
        carga_esp = pd.DataFrame(columns=["pessoa", "projeto", "mes", "nivel"])

    # ── Carrega Farol PERCEBIDO (arquivo separado) ──────────────────
    farol_manual = None
    try:
        farol_manual = pd.read_excel("Farol_Intensidade_CLEAR_2026.xlsx",
                                     sheet_name="Farol de Intensidade",
                                     header=3)
        farol_manual = farol_manual.dropna(how="all")
        if "Pessoa" in farol_manual.columns:
            farol_manual = farol_manual[farol_manual["Pessoa"].notna()]
            mask = (
                farol_manual["Pessoa"].astype(str).str.len() < 30
            ) & (
                ~farol_manual["Pessoa"].astype(str).str.contains(
                    r"TOTAL|Instrução|:", case=False, na=False, regex=True)
            )
            farol_manual = farol_manual[mask]
    except Exception:
        farol_manual = None

    # ── Toggle de modo ───────────────────────────────────────────────
    modo = st.radio(
        "Visualizar",
        ["Esperada", "Percebida", "Comparativo"],
        horizontal=True,
        captions=[
            "O que os líderes esperam",
            "Como a equipe se sente",
            "Lado a lado para identificar gaps",
        ],
    )

    # ── Helpers de cor ───────────────────────────────────────────────
    def cor_nivel(n):
        # n inteiro 0/1/2/3 (mesma escala do farol percebido)
        if n == 0: return BG_CARD, TEXTO_DIM2, ""
        if n == 1: return "#1a3d2b", "#5a9367", "🟢"
        if n == 2: return "#2d2a14", "#d99a2b", "🟡"
        return "#2d1414", "#d9534f", "🔴"

    MAP_BMA = {"B": 1, "M": 2, "A": 3}
    MES_ABREV = ["Jan","Fev","Mar","Abr","Mai","Jun",
                 "Jul","Ago","Set","Out","Nov","Dez"]

    # ── Constrói matriz ESPERADA (pessoa x mês) ──────────────────────
    # Regra: soma A=3, M=2, B=1 entre projetos; mapeia 1-2→verde, 3-4→amarelo, 5+→vermelho
    matriz_esp = {}
    if not carga_esp.empty:
        carga_esp = carga_esp.dropna(subset=["pessoa", "mes", "nivel"]).copy()
        carga_esp["mes"] = carga_esp["mes"].astype(int)
        carga_esp["peso"] = carga_esp["nivel"].astype(str).str.upper().map(MAP_BMA).fillna(0)
        soma = (carga_esp.groupby(["pessoa", "mes"])["peso"].sum()
                .reset_index().rename(columns={"peso": "total"}))
        def to_nivel(total):
            if total == 0: return 0
            if total <= 2: return 1
            if total <= 4: return 2
            return 3
        soma["nivel"] = soma["total"].apply(to_nivel)
        for _, r in soma.iterrows():
            matriz_esp[(r["pessoa"], int(r["mes"]))] = (int(r["nivel"]), int(r["total"]))

    # ── Constrói matriz PERCEBIDA (pessoa x mês) ─────────────────────
    matriz_perc = {}
    if farol_manual is not None and not farol_manual.empty and "Pessoa" in farol_manual.columns:
        for _, r in farol_manual.iterrows():
            pessoa = str(r["Pessoa"]).strip()
            for mi, col in enumerate(MES_ABREV, start=1):
                if col in farol_manual.columns:
                    v = r[col]
                    if pd.notna(v):
                        try:
                            matriz_perc[(pessoa, mi)] = int(v)
                        except Exception:
                            pass

    # Lista unificada de pessoas
    pessoas = sorted(set(
        [p for p, _ in matriz_esp.keys()] +
        [p for p, _ in matriz_perc.keys()]
    ))
    meses = list(range(1, 13))

    # Avisos de fonte
    fontes = []
    if matriz_esp:
        fontes.append(f"🎯 **Esperada**: {len(set(p for p,_ in matriz_esp.keys()))} pessoas em {len(carga_esp['projeto'].unique()) if not carga_esp.empty else 0} projetos")
    else:
        fontes.append("🎯 **Esperada**: sem dados (líderes ainda não preencheram)")
    if matriz_perc:
        fontes.append(f"👤 **Percebida**: {len(set(p for p,_ in matriz_perc.keys()))} pessoas com auto-declaração")
    else:
        fontes.append("👤 **Percebida**: sem dados (auto-declaração não preenchida)")
    st.caption(" · ".join(fontes))

    # ── Render: tabela uma coluna ou duas colunas (Comparativo) ──────
    def render_tabela(titulo, get_nivel_fn, pessoas_lista):
        col_w = 62; name_w = 130; row_h = 44
        header_cells = "".join(
            f'<th style="width:{col_w}px;min-width:{col_w}px;text-align:center;'
            f'font-size:11px;color:{TEXTO_DIM2};font-weight:600;padding:6px 2px;'
            f'border-bottom:1px solid {BORDA};">{MESES_PT[m-1]}</th>'
            for m in meses
        )
        table = (
            f'<table style="border-collapse:collapse;width:100%;font-family:sans-serif;">'
            f'<thead><tr>'
            f'<th style="width:{name_w}px;text-align:left;font-size:11px;color:{TEXTO_DIM2};'
            f'font-weight:600;padding:6px 8px;border-bottom:1px solid {BORDA};">{titulo}</th>'
            f'{header_cells}</tr></thead><tbody>'
        )
        for pi, pessoa in enumerate(pessoas_lista):
            row_bg = "rgba(255,255,255,0.03)" if pi % 2 == 1 else "transparent"
            row = f'<tr style="background:{row_bg};">'
            row += (f'<td style="padding:6px 8px;font-size:13px;font-weight:500;'
                    f'color:{TEXTO};white-space:nowrap;">{pessoa}</td>')
            for m in meses:
                nivel = get_nivel_fn(pessoa, m)
                bg, fg, icon = cor_nivel(nivel)
                row += (
                    f'<td style="text-align:center;padding:4px 2px;height:{row_h}px;">'
                    f'<div style="margin:auto;width:50px;height:34px;border-radius:6px;'
                    f'background:{bg};display:flex;flex-direction:column;'
                    f'align-items:center;justify-content:center;gap:0px;">'
                    f'<span style="font-size:15px;line-height:1;">{icon if nivel > 0 else "·"}</span>'
                    f'<span style="font-size:11px;font-weight:700;color:{fg};line-height:1.2;">'
                    f'{nivel if nivel > 0 else ""}</span></div></td>'
                )
            row += "</tr>"
            table += row
        table += "</tbody></table>"
        return table

    def n_esp(p, m):
        return matriz_esp.get((p, m), (0, 0))[0]
    def n_perc(p, m):
        return matriz_perc.get((p, m), 0)

    legenda = (
        f'<div style="display:flex;gap:20px;margin-top:12px;font-size:12px;color:{TEXTO_DIM2};">'
        f'<span>🟢 <b>1</b> — tranquilo</span>'
        f'<span>🟡 <b>2</b> — ok</span>'
        f'<span>🔴 <b>3</b> — sobrecarga</span>'
        f'<span style="opacity:.6;">· = sem dados</span>'
        f'</div>'
    )

    if modo == "Esperada":
        if not matriz_esp:
            st.info("Os líderes ainda não preencheram a aba 'Carga Esperada' nos cronogramas.")
            return
        tabela = render_tabela("Pessoa", n_esp, pessoas)
        st.markdown(
            f'<div style="background:{BG};padding:16px;border-radius:8px;border:1px solid {BORDA};'
            f'overflow-x:auto;">{tabela}{legenda}</div>',
            unsafe_allow_html=True
        )
    elif modo == "Percebida":
        if not matriz_perc:
            st.info("A equipe ainda não preencheu a auto-declaração no Farol_Intensidade.")
            return
        tabela = render_tabela("Pessoa", n_perc, pessoas)
        st.markdown(
            f'<div style="background:{BG};padding:16px;border-radius:8px;border:1px solid {BORDA};'
            f'overflow-x:auto;">{tabela}{legenda}</div>',
            unsafe_allow_html=True
        )
    else:  # Comparativo
        if not matriz_esp and not matriz_perc:
            st.info("Sem dados em nenhuma das fontes.")
            return
        col_e, col_p = st.columns(2)
        with col_e:
            st.markdown("**🎯 Esperada** (líderes)")
            if matriz_esp:
                tabela = render_tabela("Pessoa", n_esp, pessoas)
                st.markdown(
                    f'<div style="background:{BG};padding:12px;border-radius:8px;border:1px solid {BORDA};'
                    f'overflow-x:auto;">{tabela}</div>',
                    unsafe_allow_html=True
                )
            else:
                st.info("Sem dados.")
        with col_p:
            st.markdown("**👤 Percebida** (equipe)")
            if matriz_perc:
                tabela = render_tabela("Pessoa", n_perc, pessoas)
                st.markdown(
                    f'<div style="background:{BG};padding:12px;border-radius:8px;border:1px solid {BORDA};'
                    f'overflow-x:auto;">{tabela}</div>',
                    unsafe_allow_html=True
                )
            else:
                st.info("Sem dados.")
        st.markdown(legenda, unsafe_allow_html=True)

    # ── Bloco de alerta abaixo da tabela ───────────────────────────
    if modo == "Comparativo":
        # Diferença de percepção: meses em que esperada e percebida divergem
        # (considera só meses com dado real nas DUAS fontes: nível >= 1)
        difs = {}
        for pessoa in pessoas:
            linhas = []
            for m in meses:
                e = n_esp(pessoa, m)
                p = n_perc(pessoa, m)
                if e >= 1 and p >= 1 and e != p:
                    seta = "🔴↑" if p > e else "🟢↓"
                    linhas.append(f"{MESES_PT_FULL[m-1]} ({e}→{p} {seta})")
            if linhas:
                difs[pessoa] = linhas

        st.markdown("---")
        st.markdown("**🔀 Diferença de percepção** — onde o **esperado** pelos líderes e o "
                    "**percebido** pela equipe divergem:")
        if not difs:
            st.caption("Ainda não há meses com as duas fontes preenchidas para comparar. "
                       "A diferença aparece quando a pessoa tem carga esperada definida pelos "
                       "líderes e também preencheu a auto-declaração no mesmo mês.")
        else:
            for pessoa, linhas in difs.items():
                st.markdown(f"- **{pessoa}** — " + ", ".join(linhas))
            st.caption("🔴↑ a equipe sente **mais** carga que o esperado · "
                       "🟢↓ sente **menos** que o esperado · valores: esperada→percebida")
    else:
        # Picos de sobrecarga (uma linha por pessoa) — modos Esperada / Percebida
        avisos = {}
        for pessoa in pessoas:
            meses_criticos = []
            for m in meses:
                n = n_perc(pessoa, m) if modo == "Percebida" else n_esp(pessoa, m)
                if n == 3:
                    meses_criticos.append(MESES_PT_FULL[m-1])
            if meses_criticos:
                avisos[pessoa] = meses_criticos

        if avisos:
            st.markdown("---")
            rotulo = "percebida" if modo == "Percebida" else "esperada"
            st.markdown(f"**⚠️ Atenção — picos de sobrecarga ({rotulo}):**")
            for pessoa, lista_meses in avisos.items():
                meses_str = ", ".join(lista_meses)
                st.markdown(f"- **{pessoa}** — {meses_str}")


def aba_alocacao(df_gantt):
    st.subheader("Alocação — Quem está em quê, e quando")
    st.caption("Período de cada pessoa em cada frente. Fonte: alocação nova + envolvimento do master.")

    if df_gantt.empty:
        st.info("Sem dados de alocação.")
        return

    m_lo, m_hi = int(df_gantt["mes_de"].min()), int(df_gantt["mes_ate"].max())

    # ── filtros ────────────────────────────────────────────
    c1, c2 = st.columns([1, 2])
    with c1:
        vis = st.radio("Visão", ["Por pessoa", "Por projeto"], horizontal=True)
    with c2:
        if m_hi > m_lo:
            faixa = st.slider("Meses", m_lo, m_hi, (m_lo, m_hi),
                              format="%d", help="Filtra a janela de meses exibida")
        else:
            faixa = (m_lo, m_hi)

    todas_pessoas  = sorted(df_gantt["pessoa"].unique())
    todos_projetos = sorted(df_gantt["projeto"].unique())

    f1, f2 = st.columns(2)
    with f1:
        sel_pessoas = st.multiselect(
            "Filtrar pessoas", options=todas_pessoas, default=todas_pessoas,
            placeholder="Todas as pessoas",
        )
    with f2:
        sel_projetos = st.multiselect(
            "Filtrar projetos", options=todos_projetos, default=todos_projetos,
            placeholder="Todos os projetos",
        )

    m_min, m_max = faixa

    # aplica todos os filtros
    d = df_gantt[
        (df_gantt["mes_ate"] >= m_min) & (df_gantt["mes_de"] <= m_max) &
        (df_gantt["pessoa"].isin(sel_pessoas if sel_pessoas else todas_pessoas)) &
        (df_gantt["projeto"].isin(sel_projetos if sel_projetos else todos_projetos))
    ].copy()
    d["mes_de"] = d["mes_de"].clip(lower=m_min)
    d["mes_ate"] = d["mes_ate"].clip(upper=m_max)

    if d.empty:
        st.info("Nenhum resultado para os filtros selecionados.")
        return

    pessoas  = sorted(d["pessoa"].unique())
    projetos = sorted(d["projeto"].unique())
    cor_pessoa = {p: PALETA[i % len(PALETA)] for i, p in enumerate(todas_pessoas)}
    cor_proj   = {p: PALETA[i % len(PALETA)] for i, p in enumerate(todos_projetos)}

    if vis == "Por pessoa":
        groups = []
        for p in pessoas:
            dd = d[d["pessoa"] == p].sort_values("mes_de")
            groups.append((p, [(r.projeto, int(r.mes_de), int(r.mes_ate)) for r in dd.itertuples()]))
        svg, h = render_gantt(groups, cor_proj, m_min, m_max, "Pessoa")
    else:
        groups = []
        for p in projetos:
            dd = d[d["projeto"] == p].sort_values("mes_de")
            groups.append((p, [(r.pessoa, int(r.mes_de), int(r.mes_ate)) for r in dd.itertuples()]))
        svg, h = render_gantt(groups, cor_pessoa, m_min, m_max, "Projeto")

    components.html(
        f'<div style="background:{BG};margin:0;padding:0;">{svg}</div>',
        height=int(h) + 24, scrolling=False,
    )

    # Legenda de cores
    if vis == "Por pessoa":
        itens_leg = [(cor_proj.get(p, AZUL), p) for p in projetos]
        titulo_leg = "Projetos"
    else:
        itens_leg = [(cor_pessoa.get(p, AZUL), p) for p in pessoas]
        titulo_leg = "Pessoas"

    dots = "".join(
        f'<span style="display:inline-flex;align-items:center;gap:5px;'
        f'margin:3px 10px 3px 0;font-size:12px;color:{TEXTO_DIM};">'
        f'<span style="width:12px;height:12px;border-radius:3px;background:{cor};'
        f'flex-shrink:0;display:inline-block;"></span>{nome}</span>'
        for cor, nome in itens_leg
    )
    st.markdown(
        f'<div style="margin-top:8px;padding:10px 14px;background:{BG_CARD};'
        f'border:1px solid {BORDA};border-radius:8px;">'
        f'<span style="font-size:11px;font-weight:600;color:{TEXTO_DIM2};'
        f'text-transform:uppercase;letter-spacing:.05em;">{titulo_leg}</span><br>'
        f'<div style="margin-top:6px;">{dots}</div></div>',
        unsafe_allow_html=True
    )

    # carga: pico de frentes simultâneas por pessoa
    st.markdown("**Pico de frentes simultâneas (no período exibido)**")
    pico = {}
    for p in pessoas:
        dd = d[d["pessoa"] == p]
        pk = max((int(((dd["mes_de"] <= m) & (dd["mes_ate"] >= m)).sum())
                  for m in range(m_min, m_max + 1)), default=0)
        pico[p] = pk
    cols = st.columns(min(len(pessoas), 7) or 1)
    for i, p in enumerate(sorted(pico, key=lambda x: -pico[x])):
        cols[i % len(cols)].metric(p, pico[p])


# ============================================================
# ABA — Comunicação (painel dedicado)
# ============================================================
def aba_comunicacao(df_at_full):
    st.subheader("Comunicação")
    st.caption("Painel dedicado do projeto de Comunicação — mostra todas as "
               "atividades, independente do filtro lateral.")

    com = df_at_full[df_at_full["projeto"] == "Comunicação"].copy()
    if com.empty:
        st.info("Sem atividades de Comunicação no master.")
        return

    com_data = com[com["prazo"].notna()].copy()
    com_data = com_data[(com_data["prazo"] >= pd.Timestamp("2026-01-01")) &
                        (com_data["prazo"] <= pd.Timestamp("2026-12-31"))]
    sem_data = com[com["prazo"].isna()].copy()

    st.markdown(
        f'<div style="border-left:4px solid {COR_COMUNICACAO}; background:{BG_CARD}; '
        f'padding:8px 12px; border-radius:4px; margin-bottom:12px; color:{TEXTO_DIM}; '
        f'font-size:0.8rem;">Identidade visual do projeto: '
        f'<span style="display:inline-block;width:11px;height:11px;background:{COR_COMUNICACAO};'
        f'border-radius:2px;margin:0 4px -1px 4px;"></span>âmbar — a mesma cor usada no Calendário.'
        f'</div>',
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Atividades", len(com))
    c2.metric("Entregáveis", int(com["eh_entregavel"].sum()))
    c3.metric("Com data", len(com_data))
    c4.metric("Sem data", len(sem_data))

    st.markdown("---")
    st.markdown("**Linha do tempo** — todas as atividades datadas (◆ = entregável crítico)")

    if com_data.empty:
        st.info("Nenhuma atividade de Comunicação com data definida.")
    else:
        com_data = com_data.sort_values("prazo")
        com_data["_mes"] = com_data["prazo"].dt.month
        for mes in sorted(com_data["_mes"].unique()):
            bloco = com_data[com_data["_mes"] == mes]
            st.markdown(
                f'<div style="font-size:0.8rem; font-weight:600; color:{COR_COMUNICACAO}; '
                f'margin:14px 0 6px 0;">{MESES_PT_FULL[mes - 1]} 2026 · '
                f'{len(bloco)} atividade(s)</div>',
                unsafe_allow_html=True,
            )
            linhas = ""
            for _, r in bloco.iterrows():
                marca = "◆ " if r["eh_entregavel"] else ""
                resp = (str(r["responsaveis"]) if pd.notna(r["responsaveis"])
                        and str(r["responsaveis"]).strip() else "—")
                sub = (f' · {r["sub_projeto"]}' if pd.notna(r["sub_projeto"])
                       and str(r["sub_projeto"]).strip() else "")
                cor_st = CORES_STATUS.get(r["status"], TEXTO_DIM)
                linhas += (
                    f'<div style="display:flex; align-items:center; gap:10px; '
                    f'background:{BG_CARD}; border-left:3px solid {COR_COMUNICACAO}; '
                    f'padding:6px 10px; margin-bottom:4px; border-radius:3px;">'
                    f'<div style="min-width:50px; font-size:0.72rem; color:{TEXTO_DIM2}; '
                    f'font-weight:600;">{r["prazo"].strftime("%d/%m")}</div>'
                    f'<div style="flex:1; font-size:0.82rem; color:{TEXTO};">{marca}'
                    f'{r["atividade"]}<span style="color:{TEXTO_DIM2}; '
                    f'font-size:0.72rem;">{sub}</span></div>'
                    f'<div style="font-size:0.72rem; color:{TEXTO_DIM}; min-width:90px; '
                    f'text-align:right;">{resp}</div>'
                    f'<div style="font-size:0.68rem; color:{cor_st}; min-width:80px; '
                    f'text-align:right;">{r["status"]}</div></div>'
                )
            st.markdown(linhas, unsafe_allow_html=True)

    if not sem_data.empty:
        st.markdown("---")
        with st.expander(f"Atividades sem data definida ({len(sem_data)})"):
            for _, r in sem_data.iterrows():
                marca = "◆ " if r["eh_entregavel"] else ""
                obs = (f' — {r["prazo_obs"]}' if pd.notna(r["prazo_obs"])
                       and str(r["prazo_obs"]).strip() else "")
                st.markdown(f'- {marca}{r["atividade"]}{obs}')


# ============================================================
# ABA — Férias (linha do tempo, para ver sobreposições)
# ============================================================
def aba_ferias():
    st.markdown("### Férias da equipe")

    # download do modelo
    try:
        with open(FERIAS_FILE, "rb") as fh:
            st.download_button("Baixar modelo de férias", fh.read(),
                               file_name="Ferias_CLEAR_2026.xlsx",
                               mime=("application/vnd.openxmlformats-"
                                     "officedocument.spreadsheetml.sheet"))
    except Exception:
        pass

    try:
        df = pd.read_excel(FERIAS_FILE, sheet_name="Férias", skiprows=2)
    except Exception:
        st.info("Arquivo de férias ainda não disponível no repositório.")
        return

    df = df.rename(columns=lambda c: str(c).strip())
    if "Pessoa" not in df.columns:
        st.info("Planilha de férias sem a coluna 'Pessoa'.")
        return

    CANON = {"Pleno": "Fred", "Senior": "Hisrael", "Sênior": "Hisrael",
             "Luiggi": "Luigi", "BiaS": "Bia S"}
    df["Pessoa"] = df["Pessoa"].map(lambda x: CANON.get(str(x).strip(), str(x).strip()))
    df["Início"] = pd.to_datetime(df.get("Início"), errors="coerce")
    df["Fim"] = pd.to_datetime(df.get("Fim"), errors="coerce")
    df = df.dropna(subset=["Início", "Fim"])
    df = df[df["Pessoa"].isin(EQUIPE)]

    if df.empty:
        st.info("Ninguém preencheu as férias ainda. Baixe o modelo acima, "
                "preencha Início/Fim e suba o arquivo no GitHub.")
        return

    df["Obs"] = df.get("Observação", "").fillna("") if "Observação" in df.columns else ""
    # px.timeline é fim-exclusivo: soma 1 dia para o bloco cobrir o dia final
    df["Fim_plot"] = df["Fim"] + pd.Timedelta(days=1)

    ordem = [p for p in EQUIPE if p in set(df["Pessoa"])]
    cmap = {p: PALETA[i % len(PALETA)] for i, p in enumerate(EQUIPE)}

    c1, c2, c3 = st.columns(3)
    c1.metric("Pessoas com férias", df["Pessoa"].nunique())
    c2.metric("Períodos", len(df))
    dias_tot = int(((df["Fim"] - df["Início"]).dt.days + 1).sum())
    c3.metric("Dias-pessoa", dias_tot)

    fig = px.timeline(
        df, x_start="Início", x_end="Fim_plot", y="Pessoa",
        color="Pessoa", color_discrete_map=cmap,
        custom_data=["Início", "Fim", "Obs"],
    )
    fig.update_traces(hovertemplate=(
        "<b>%{y}</b><br>%{customdata[0]|%d/%m/%Y} – "
        "%{customdata[1]|%d/%m/%Y}<br>%{customdata[2]}<extra></extra>"))
    fig.update_yaxes(categoryorder="array", categoryarray=list(reversed(ordem)),
                     title=None, gridcolor=BORDA)
    fig.update_xaxes(gridcolor=BORDA, title=None, dtick="M1", tickformat="%b/%y")
    fig.update_layout(
        height=90 + 34 * len(ordem), margin=dict(t=10, b=10, l=10, r=10),
        paper_bgcolor=BG, plot_bgcolor=BG, font=dict(color=TEXTO_DIM),
        showlegend=False, bargap=0.35,
    )

    # sombrear janelas com 2+ pessoas de férias simultâneas
    dias = pd.date_range(df["Início"].min(), df["Fim"].max(), freq="D")
    por_dia = {d: [p for p, s, e in zip(df["Pessoa"], df["Início"], df["Fim"])
                   if s <= d <= e] for d in dias}
    janelas, cur = [], None
    for d in dias:
        quem = tuple(sorted(set(por_dia[d])))
        if len(quem) >= 2:
            if cur and cur["quem"] == quem:
                cur["fim"] = d
            else:
                if cur:
                    janelas.append(cur)
                cur = {"ini": d, "fim": d, "quem": quem}
        elif cur:
            janelas.append(cur); cur = None
    if cur:
        janelas.append(cur)

    for j in janelas:
        fig.add_vrect(x0=j["ini"], x1=j["fim"] + pd.Timedelta(days=1),
                      fillcolor=AZUL_CLARO, opacity=0.10, line_width=0, layer="below")

    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    if janelas:
        st.markdown("**Sobreposições**")
        for j in janelas:
            ini, fim = j["ini"].strftime("%d/%m"), j["fim"].strftime("%d/%m/%Y")
            st.markdown(f'- {ini}–{fim} · {", ".join(j["quem"])}')
    else:
        st.caption("Sem sobreposições entre os períodos marcados.")


# ============================================================
# Main
# ============================================================
def main():
    try:
        df_at, df_resp, df_aloc, df_pessoas, df_gantt = carregar_dados()
    except FileNotFoundError:
        st.error("Planilha mestra não encontrada.")
        st.stop()

    escopo, projs = sidebar_filtros(df_at)
    df_at_f = filtrar(df_at, escopo, projs)
    df_resp_f = df_resp[df_resp["projeto"].isin(projs)].copy()
    if escopo == "Só entregáveis críticos":
        df_resp_f = df_resp_f[df_resp_f["eh_entregavel"]]

    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs(
        ["Visão Geral", "Calendário", "Comunicação", "Alocação",
         "Equipe", "Pesquisador", "Farol", "Férias"])
    with tab1:
        aba_gantt(df_at_f)
    with tab2:
        aba_calendario(df_at_f)
    with tab3:
        aba_comunicacao(df_at)
    with tab4:
        aba_alocacao(df_gantt)
    with tab5:
        aba_equipe(df_aloc)
    with tab6:
        aba_pesquisador(df_resp_f, df_aloc, df_pessoas)
    with tab7:
        aba_farol(df_aloc)
    with tab8:
        aba_ferias()


if __name__ == "__main__":
    main()
