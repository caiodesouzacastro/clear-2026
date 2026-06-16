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
# ABA 4 — Farol de Intensidade de Pessoal
# ============================================================
def aba_farol(df_aloc):
    st.subheader("Farol de Intensidade — Carga por Pessoa")
    st.caption("Percepção de carga de trabalho por pessoa, mês a mês. "
               "1 = tranquilo · 2 = ok · 3 = sobrecarga.")

    if df_aloc.empty:
        st.info("Sem dados de alocação.")
        return

    # ── Tenta ler o Farol_Intensidade preenchido (auto-declaração) ───
    farol_manual = None
    try:
        farol_manual = pd.read_excel("Farol_Intensidade_CLEAR_2026.xlsx",
                                     sheet_name="Farol de Intensidade",
                                     header=3)
        # remove colunas vazias e a linha de "TOTAL EQUIPE"
        farol_manual = farol_manual.dropna(how="all")
        if "Pessoa" in farol_manual.columns:
            farol_manual = farol_manual[farol_manual["Pessoa"].notna()]
            farol_manual = farol_manual[~farol_manual["Pessoa"].astype(str)
                                        .str.contains("TOTAL", case=False, na=False)]
    except Exception:
        farol_manual = None

    # ── Caso contrário, calcula da alocação automática (escala 1-3) ──
    MES_NUM = {m: i + 1 for i, m in enumerate(MESES_PT_FULL)}
    aloc = df_aloc.copy()
    aloc["mnum"] = aloc["mes"].map(MES_NUM)
    aloc = aloc.dropna(subset=["mnum", "pessoa"])
    aloc = aloc[aloc["pessoa"].notna() & (aloc["pessoa"] != "PMO")]

    pivot = (aloc.groupby(["pessoa", "mnum"])["projeto_alocacao"]
             .nunique().reset_index()
             .rename(columns={"projeto_alocacao": "n_proj"}))

    # Mapeia número de projetos -> escala 1-3 (heurística)
    def to_escala(n):
        if n == 0: return 0
        if n == 1: return 1
        if n <= 3: return 2
        return 3

    pivot["nivel"] = pivot["n_proj"].apply(to_escala)
    pessoas = sorted(pivot["pessoa"].unique())
    meses = list(range(1, 13))

    fonte_label = "Auto-calculado a partir da alocação"
    if farol_manual is not None and not farol_manual.empty:
        fonte_label = "Preenchido pela equipe (auto-declaração)"

    st.caption(f"📊 Fonte: {fonte_label}")

    # Cor por nível (1-3)
    def cor_nivel(nivel):
        if nivel == 0: return BG_CARD, TEXTO_DIM2, ""
        if nivel == 1: return "#1a3d2b", "#5a9367", "🟢"  # tranquilo
        if nivel == 2: return "#2d2a14", "#d99a2b", "🟡"  # ok
        return "#2d1414", "#d9534f", "🔴"                  # sobrecarga

    # Tabela HTML
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
        f'font-weight:600;padding:6px 8px;border-bottom:1px solid {BORDA};">Pessoa</th>'
        f'{header_cells}</tr></thead><tbody>'
    )

    # Função pra pegar o nível de uma pessoa em um mês
    MES_ABREV = ["Jan","Fev","Mar","Abr","Mai","Jun",
                 "Jul","Ago","Set","Out","Nov","Dez"]
    def nivel_de(pessoa, m):
        if farol_manual is not None and "Pessoa" in farol_manual.columns:
            linha = farol_manual[farol_manual["Pessoa"] == pessoa]
            if not linha.empty:
                col = MES_ABREV[m-1]
                if col in linha.columns:
                    v = linha[col].iloc[0]
                    if pd.notna(v):
                        try: return int(v)
                        except: return 0
            return 0
        sub = pivot[(pivot["pessoa"] == pessoa) & (pivot["mnum"] == m)]
        return int(sub["nivel"].iloc[0]) if not sub.empty else 0

    # Lista canônica de pessoas (do manual se houver, senão do auto)
    if farol_manual is not None and not farol_manual.empty and "Pessoa" in farol_manual.columns:
        pessoas = sorted(farol_manual["Pessoa"].dropna().unique())

    for pi, pessoa in enumerate(pessoas):
        row_bg = "rgba(255,255,255,0.03)" if pi % 2 == 1 else "transparent"
        row = f'<tr style="background:{row_bg};">'
        row += (f'<td style="padding:6px 8px;font-size:13px;font-weight:500;'
                f'color:{TEXTO};white-space:nowrap;">{pessoa}</td>')
        for m in meses:
            nivel = nivel_de(pessoa, m)
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

    legenda = (
        f'<div style="display:flex;gap:20px;margin-top:12px;font-size:12px;color:{TEXTO_DIM2};">'
        f'<span>🟢 <b>1</b> — tranquilo</span>'
        f'<span>🟡 <b>2</b> — ok</span>'
        f'<span>🔴 <b>3</b> — sobrecarga</span>'
        f'<span style="opacity:.6;">· = sem dados</span>'
        f'</div>'
    )

    st.markdown(
        f'<div style="background:{BG};padding:16px;border-radius:8px;border:1px solid {BORDA};'
        f'overflow-x:auto;">{table}{legenda}</div>',
        unsafe_allow_html=True
    )

    # Destaque: quem está em sobrecarga (nível 3) em algum mês
    avisos = []
    for pessoa in pessoas:
        for m in meses:
            if nivel_de(pessoa, m) == 3:
                avisos.append((pessoa, MESES_PT_FULL[m-1]))

    if avisos:
        st.markdown("---")
        st.markdown("**⚠️ Atenção — picos de sobrecarga:**")
        for pessoa, mes in avisos:
            st.markdown(f"- **{pessoa}** em {mes}")


# ============================================================
# ABA 5 — Alocação (Gantt por pessoa × projeto)
# ============================================================
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

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        ["Visão Geral", "Calendário", "Alocação", "Farol", "Equipe", "Pesquisador"])
    with tab1:
        aba_gantt(df_at_f)
    with tab2:
        aba_calendario(df_at_f)
    with tab3:
        aba_alocacao(df_gantt)
    with tab4:
        aba_farol(df_aloc)
    with tab5:
        aba_equipe(df_aloc)
    with tab6:
        aba_pesquisador(df_resp_f, df_aloc, df_pessoas)


if __name__ == "__main__":
    main()
