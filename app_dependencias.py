"""
CLEAR 2026 — Cadeia de Dependências (v3, timeline conectada)

Formato: cada entregável é um ponto num eixo de tempo (posicionado pelo prazo),
agrupado por projeto na vertical. As arestas viram linhas curvas.
Nós sem data ficam empilhados à direita numa "faixa sem data".

Quatro abas:
  1. Impacto a jusante  (timeline conectada; escolhe um nó → cadeia downstream)
  2. Caminho crítico    (mesmo eixo; ilumina a corrente mais longa)
  3. Todas as atividades (as 736 do Master, Gantt por projeto)
  4. Inter-projetos & flags
"""

import collections
import datetime as dt
from pathlib import Path

import streamlit as st
import openpyxl
import plotly.express as px
import plotly.graph_objects as go

BASE = Path(__file__).parent
MASTER = BASE / "CLEAR_Master_2026.xlsx"
DEPS = BASE / "Dependencias_CLEAR_2026.xlsx"

# -------------------------------------------------------------------- paleta
BG        = "#0A1929"
PANEL     = "#132F4C"
PANEL2    = "#1E3A5F"
PRIMARY   = "#5090D3"
TEXT      = "#FFFFFF"
MUTED     = "#B2BAC2"
BORDER    = "#26456B"

# cores de status com MAIS contraste do que o app original, p/ o grafo ficar legível
# (mantém o mesmo idioma cromático, mas espaça os tons)
STATUS_COR = {
    "concluído":         "#5DCAA5",   # verde-água — se destaca sem virar semáforo
    "pronto":            "#5DCAA5",
    "enviado para eesp": "#5DCAA5",
    "em andamento":      "#5090D3",   # azul primário
    "não iniciado":      "#6F7E8C",   # cinza (não é o foco)
    "atrasado":          "#E57373",   # vermelho — agora atrasado grita
    "reunião":           "#B39DDB",
}
COR_PLACEHOLDER = "#D4A55E"
COR_DESCONHECIDO = "#6F7E8C"
COR_FADE = "#26456B"
COR_LINHA_ARESTA = "#3D6A9E"
COR_LINHA_DESTAQUE = "#7BB3F0"

DERIVAVEIS = {"Painel CLEAR", "EVALAC"}


def _status_norm(s):
    if not s: return ""
    return str(s).strip()


def _cor_status(status_txt, fonte=None):
    if fonte == "placeholder":
        return COR_PLACEHOLDER
    if not status_txt:
        return COR_DESCONHECIDO
    return STATUS_COR.get(status_txt.strip().lower(), COR_DESCONHECIDO)


# ============================================================ auth
def _senha_ok() -> bool:
    if st.session_state.get("_auth"):
        return True
    pw = st.secrets.get("app_password") if hasattr(st, "secrets") else None
    if not pw:
        return True
    val = st.text_input("Senha", type="password")
    if val == pw:
        st.session_state["_auth"] = True
        st.rerun()
    if val:
        st.error("Senha incorreta.")
    return False


# ============================================================ dados
@st.cache_data(show_spinner=False)
def carregar():
    wm = openpyxl.load_workbook(MASTER, read_only=True, data_only=True)
    ws = wm["Atividades"]
    hdr = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    H = {h: i for i, h in enumerate(hdr)}
    master = {}
    atividades = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        proj, ativ, det = r[H["projeto"]], r[H["atividade"]], r[H["detalhe"]]
        info = {"status": r[H["status"]], "prazo": r[H["prazo"]],
                "sub": r[H["sub_projeto"]], "eh_ent": r[H["eh_entregavel"]] is True,
                "resp": r[H["responsaveis"]]}
        master[(proj, ativ, det)] = info
        atividades.append({
            "projeto": proj or "(sem projeto)", "sub": info["sub"] or "",
            "atividade": ativ or "", "detalhe": det or "",
            "prazo": info["prazo"], "status": info["status"] or "Não Iniciado",
            "responsaveis": info["resp"] or "", "eh_entregavel": info["eh_ent"],
        })

    wd = openpyxl.load_workbook(DEPS, read_only=True, data_only=True)
    nos = {}
    for r in wd["Nos"].iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        did, proj, ativ, det, prazo_dep, _, fonte = r[0], r[1], r[2], r[3], r[4], r[5], r[6]
        m = master.get((proj, ativ, det))
        nos[did] = {
            "projeto": proj, "atividade": ativ, "detalhe": det,
            "fonte": fonte,
            "orfao": (fonte != "placeholder") and (m is None),
            "status": m["status"] if m else None,
            "prazo": m["prazo"] if m else prazo_dep,
        }
    edges = []
    for r in wd["Arestas"].iter_rows(min_row=2, values_only=True):
        if not r[0] or not r[1]:
            continue
        edges.append({"o": r[0], "d": r[1],
                      "tipo": r[2] or "insumo", "forca": r[3] or "dura",
                      "nota": r[5] or ""})

    derivadas = []
    for proj in DERIVAVEIS:
        ns = [d for d, n in nos.items() if n["projeto"] == proj and not n["orfao"]]
        ns.sort(key=lambda d: (nos[d]["prazo"] is None, str(nos[d]["prazo"])))
        derivadas += [{"o": a, "d": b} for a, b in zip(ns, ns[1:])]

    flags = []
    if "Flags" in wd.sheetnames:
        for r in wd["Flags"].iter_rows(min_row=2, values_only=True):
            if r[0]:
                flags.append({"item": r[0], "tipo": r[1] or "", "nota": r[2] or ""})

    return nos, edges, derivadas, atividades, flags


# ============================================================ grafo (só p/ álgebra)
def _grafo(edges):
    adj, radj = collections.defaultdict(list), collections.defaultdict(list)
    for e in edges:
        adj[e["o"]].append(e["d"]); radj[e["d"]].append(e["o"])
    return adj, radj


def _downstream(adj, s):
    seen, st = set(), [s]
    while st:
        u = st.pop()
        for v in adj[u]:
            if v not in seen:
                seen.add(v); st.append(v)
    return seen


def _caminho_mais_longo(radj, t):
    memo = {}
    def dfs(u):
        if u in memo: return memo[u]
        if not radj[u]:
            memo[u] = [u]; return memo[u]
        memo[u] = max((dfs(p) for p in radj[u]), key=len) + [u]
        return memo[u]
    return dfs(t)


# ============================================================ TIMELINE conectada
def _prazo_dt(p):
    """Normaliza prazo pra datetime (aceita date, datetime, string)."""
    if p is None: return None
    if isinstance(p, dt.datetime): return p
    if isinstance(p, dt.date): return dt.datetime(p.year, p.month, p.day)
    try: return dt.datetime.fromisoformat(str(p)[:10])
    except Exception: return None


def _timeline_conectada(nos, edges, *, destaque_nos=None, destaque_arestas=None,
                        titulo=""):
    """
    Cada nó = ponto (x=prazo, y=projeto). Nós sem data ficam empilhados numa
    faixa 'sem data' à direita. Arestas viram linhas curvas.
    """
    destaque_nos = destaque_nos or set()
    destaque_arestas = destaque_arestas or set()
    tem_destaque = bool(destaque_nos)

    # projetos: ordena por prazo mediano (projetos "cedo" no topo)
    por_proj = collections.defaultdict(list)
    for d, n in nos.items():
        por_proj[n["projeto"] or "(sem projeto)"].append(d)
    def med_prazo(dids):
        ds = sorted([_prazo_dt(nos[d]["prazo"]) for d in dids
                     if _prazo_dt(nos[d]["prazo"])])
        return ds[len(ds)//2] if ds else dt.datetime(2099, 1, 1)
    proj_ordem = sorted(por_proj, key=lambda p: med_prazo(por_proj[p]))
    proj_y = {p: i for i, p in enumerate(proj_ordem)}

    # eixo x: min/max de prazos reais + coluna extra p/ "sem data"
    todas_datas = [d for d in (_prazo_dt(n["prazo"]) for n in nos.values()) if d]
    if todas_datas:
        dmin, dmax = min(todas_datas), max(todas_datas)
        span = (dmax - dmin).days or 30
        margem = dt.timedelta(days=max(int(span * 0.05), 4))
        x_sem_data = dmax + dt.timedelta(days=max(int(span * 0.12), 10))
    else:
        dmin = dt.datetime.today(); dmax = dmin
        margem = dt.timedelta(days=5); x_sem_data = dmin + dt.timedelta(days=15)

    # posições dos nós; empilha "sem data" com pequeno jitter vertical
    pos = {}
    empilhados = collections.Counter()
    for did, n in nos.items():
        y = proj_y[n["projeto"] or "(sem projeto)"]
        p = _prazo_dt(n["prazo"])
        if p is None:
            k = empilhados[y]; empilhados[y] += 1
            pos[did] = (x_sem_data + dt.timedelta(days=k * 3), y + 0.06 * k)
        else:
            pos[did] = (p, y)

    fig = go.Figure()

    # ------ arestas: linha curva (quadrática) via 20 pontos
    def curva(x0, y0, x1, y1, bend=0.35):
        # ponto de controle no meio, com desvio proporcional à distância vertical
        # se dy≠0 encurva pra fora; se dy=0 encurva um pouco pra baixo
        dy = y1 - y0
        cx = x0 + (x1 - x0) / 2
        cy = (y0 + y1) / 2 + (bend if dy == 0 else 0)
        # bezier quadrática amostrada
        xs, ys = [], []
        for i in range(21):
            t = i / 20
            xt = (1 - t) ** 2 * (x0.timestamp() if hasattr(x0, "timestamp") else x0) \
                 + 2 * (1 - t) * t * (cx.timestamp() if hasattr(cx, "timestamp") else cx) \
                 + t ** 2 * (x1.timestamp() if hasattr(x1, "timestamp") else x1)
            yt = (1 - t) ** 2 * y0 + 2 * (1 - t) * t * cy + t ** 2 * y1
            xs.append(dt.datetime.fromtimestamp(xt))
            ys.append(yt)
        return xs, ys

    def _placeholder(did):
        return nos[did]["fonte"] == "placeholder"

    for e in edges:
        if e["o"] not in pos or e["d"] not in pos:
            continue
        (x0, y0), (x1, y1) = pos[e["o"]], pos[e["d"]]
        destacada = (e["o"], e["d"]) in destaque_arestas
        pontilhada = _placeholder(e["o"]) or _placeholder(e["d"]) \
                     or nos[e["o"]]["prazo"] is None or nos[e["d"]]["prazo"] is None
        if destacada:
            cor, w = COR_LINHA_DESTAQUE, 3
        elif tem_destaque:
            cor, w = BORDER, 1
        else:
            cor, w = COR_LINHA_ARESTA, 1.5
        xs, ys = curva(x0, y0, x1, y1)
        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode="lines",
            line=dict(color=cor, width=w, dash="dot" if pontilhada else "solid"),
            hoverinfo="skip", showlegend=False,
        ))
        # ponta da seta (triângulo pequeno perto do destino)
        fig.add_trace(go.Scatter(
            x=[xs[-2], xs[-1]], y=[ys[-2], ys[-1]], mode="lines",
            line=dict(color=cor, width=w + 1),
            hoverinfo="skip", showlegend=False,
        ))

    # ------ nós: um scatter por status pra ter legenda
    grupos = collections.defaultdict(list)
    for did, n in nos.items():
        chave = ("PLACEHOLDER" if n["fonte"] == "placeholder"
                 else _status_norm(n["status"]) or "Não Iniciado")
        grupos[chave].append(did)

    for chave, dids in grupos.items():
        xs, ys, texts, hovers, cores, bordas, tamanhos = [], [], [], [], [], [], []
        for did in dids:
            x, y = pos[did]
            n = nos[did]
            fade = tem_destaque and did not in destaque_nos
            cor = COR_FADE if fade else _cor_status(n["status"], n["fonte"])
            borda = BORDER if fade else (PRIMARY if did in destaque_nos else BG)
            tamanho = 24 if did in destaque_nos else 18
            xs.append(x); ys.append(y)
            nome = (n["atividade"] or "")[:38]
            texts.append(f"<b>{nome}</b>")
            prazo_txt = (_prazo_dt(n["prazo"]).date().isoformat()
                         if _prazo_dt(n["prazo"]) else "sem data")
            hovers.append(
                f"<b>{n['atividade']}</b><br>"
                f"{n['projeto']}<br>"
                f"prazo: {prazo_txt}<br>"
                f"status: {n['status'] or '—'}<br>"
                f"<i>{did}</i>"
            )
            cores.append(cor); bordas.append(borda); tamanhos.append(tamanho)
        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode="markers+text",
            marker=dict(size=tamanhos, color=cores,
                        line=dict(color=bordas, width=2)),
            text=texts, textposition="middle right",
            textfont=dict(color=TEXT, size=11, family="Arial"),
            hovertext=hovers, hoverinfo="text",
            name=chave, showlegend=True,
        ))

    # linha vertical separando "com data" de "sem data"
    if any(nos[d]["prazo"] is None for d in nos):
        fig.add_vline(x=(dmax + margem), line_dash="dot",
                      line_color=BORDER, line_width=1)
        fig.add_annotation(x=x_sem_data, y=-0.7, text="sem data",
                           showarrow=False, font=dict(color=MUTED, size=10))

    n_proj = len(proj_ordem)
    fig.update_layout(
        title=dict(text=titulo, font=dict(color=TEXT, size=14), x=0.01),
        plot_bgcolor=BG, paper_bgcolor=BG,
        font=dict(color=TEXT, family="Arial", size=11),
        legend=dict(bgcolor=PANEL, bordercolor=BORDER, borderwidth=1,
                    orientation="h", yanchor="bottom", y=1.02, x=0),
        margin=dict(l=8, r=200, t=40, b=8),
        height=max(420, 90 * n_proj + 120),
        xaxis=dict(
            gridcolor=PANEL2, zerolinecolor=PANEL2,
            range=[dmin - margem, x_sem_data + dt.timedelta(days=8)],
            tickformat="%b/%y",
        ),
        yaxis=dict(
            gridcolor=PANEL2, zerolinecolor=PANEL2,
            tickmode="array",
            tickvals=list(proj_y.values()),
            ticktext=proj_ordem,
            autorange="reversed",
        ),
        hoverlabel=dict(bgcolor=PANEL, bordercolor=PRIMARY,
                        font=dict(color=TEXT, family="Arial")),
    )
    return fig


# ============================================================ UI
def main():
    st.set_page_config(page_title="CLEAR 2026 · Dependências",
                       page_icon="🔗", layout="wide",
                       initial_sidebar_state="collapsed")

    st.markdown(f"""
    <style>
      .block-container {{ padding-top: 1.2rem; padding-bottom: 1rem; max-width: 100%; }}
      h1, h2, h3 {{ color: {TEXT}; }}
      .stTabs [data-baseweb="tab-list"] {{ gap: 4px; }}
      .stTabs [data-baseweb="tab"] {{
        background: {PANEL}; color: {MUTED}; border-radius: 8px 8px 0 0;
        padding: 8px 16px; font-weight: 500;
      }}
      .stTabs [aria-selected="true"] {{ background: {PANEL2}; color: {TEXT}; }}
    </style>
    """, unsafe_allow_html=True)

    st.title("🔗 Cadeia de Dependências — CLEAR 2026")
    st.caption("Cada ponto = um entregável posicionado pela sua data de entrega. "
               "Linhas conectam **quem destrava quem** (esquerda destrava direita). "
               "Linhas pontilhadas = alguma ponta ainda sem data.")

    if not _senha_ok():
        st.stop()

    try:
        nos, edges, derivadas, atividades, flags = carregar()
    except FileNotFoundError as e:
        st.error(f"Arquivo não encontrado: {e.filename}"); st.stop()

    orfaos = [d for d, n in nos.items() if n["orfao"]]
    if orfaos:
        st.warning(f"⚠️ Nós órfãos: {', '.join(orfaos)}")

    adj, radj = _grafo(edges)

    def rotulo(d):
        return f"{nos[d]['atividade'][:52]}  ({d})"

    t1, t2, t3, t4 = st.tabs([
        "⬇️ Impacto a jusante",
        "🎯 Caminho crítico",
        "📅 Todas as atividades",
        "🔀 Inter-projetos & flags",
    ])

    # ---------------------------------------------------------- Tab 1
    with t1:
        col1, col2 = st.columns([3, 1])
        candidatos = sorted(d for d in nos if adj[d] or radj[d])
        with col1:
            alvo = st.selectbox("Se esta atividade atrasar, o que cai junto?",
                                candidatos, format_func=rotulo, key="t1_sel")
        down = _downstream(adj, alvo)
        destaque = {alvo} | down
        arestas_dest = {(e["o"], e["d"]) for e in edges
                        if e["o"] in destaque and e["d"] in destaque}
        chegadas = sorted(d for d in down if not adj[d])
        with col2:
            st.metric("Atividades afetadas", len(down))
            if chegadas:
                st.markdown("**Marcos atingidos:**")
                for d in chegadas:
                    st.markdown(f"- {nos[d]['atividade'][:36]}")
        st.plotly_chart(
            _timeline_conectada(nos, edges,
                                destaque_nos=destaque,
                                destaque_arestas=arestas_dest),
            use_container_width=True, config={"displayModeBar": False},
        )

    # ---------------------------------------------------------- Tab 2
    with t2:
        chegadas = sorted(d for d in nos if not adj[d] and radj[d])
        alvo = st.selectbox("Marco final:", chegadas,
                            format_func=rotulo, key="t2_sel")
        caminho = _caminho_mais_longo(radj, alvo)
        destaque = set(caminho)
        arestas_dest = set(zip(caminho, caminho[1:]))
        st.markdown("**Corrente crítica:** " + " → ".join(
            f"*{nos[c]['atividade'][:26]}*" for c in caminho))
        st.plotly_chart(
            _timeline_conectada(nos, edges,
                                destaque_nos=destaque,
                                destaque_arestas=arestas_dest),
            use_container_width=True, config={"displayModeBar": False},
        )

    # ---------------------------------------------------------- Tab 3
    with t3:
        st.markdown("**Timeline de todas as atividades do Master** — "
                    "cada barra = uma atividade, colorida por status.")
        projs = sorted({a["projeto"] for a in atividades if a["projeto"]})
        f1, f2, f3 = st.columns([2, 2, 1])
        with f1: sel_projs = st.multiselect("Projeto(s)", projs, default=projs)
        with f2:
            sel_status = st.multiselect(
                "Status",
                ["Concluído", "Em Andamento", "Não Iniciado", "Atrasado", "Reunião"],
                default=["Em Andamento", "Não Iniciado", "Atrasado"],
            )
        with f3: so_ent = st.checkbox("Só entregáveis", value=False)

        filt = [a for a in atividades
                if a["projeto"] in sel_projs
                and _prazo_dt(a["prazo"]) is not None
                and str(a["status"]).strip() in sel_status
                and (not so_ent or a["eh_entregavel"])]
        sem_data = sum(1 for a in atividades
                       if a["projeto"] in sel_projs and _prazo_dt(a["prazo"]) is None
                       and (not so_ent or a["eh_entregavel"]))
        st.caption(f"{len(filt)} atividades · {sem_data} sem data (ocultas)")

        if not filt:
            st.info("Nada a mostrar com esse filtro.")
        else:
            for a in filt:
                p = _prazo_dt(a["prazo"])
                a["_start"] = p; a["_end"] = p + dt.timedelta(days=1)
                a["_status_norm"] = str(a["status"]).strip()
            fig = px.timeline(
                filt, x_start="_start", x_end="_end", y="projeto",
                color="_status_norm",
                color_discrete_map={k.capitalize(): v for k, v in STATUS_COR.items()
                                    if k in ("concluído", "em andamento",
                                             "não iniciado", "atrasado", "reunião")},
                hover_data={"atividade": True, "responsaveis": True, "sub": True,
                            "_start": False, "_end": False, "projeto": False,
                            "_status_norm": True},
                height=max(360, 34 * len(sel_projs) + 120),
            )
            ent = [a for a in filt if a["eh_entregavel"]]
            if ent:
                fig.add_trace(go.Scatter(
                    x=[_prazo_dt(a["prazo"]) for a in ent],
                    y=[a["projeto"] for a in ent],
                    mode="markers",
                    marker=dict(symbol="diamond", size=11, color="#FFFFFF",
                                line=dict(color=PRIMARY, width=1.5)),
                    name="Entregável",
                    hovertext=[a["atividade"][:80] for a in ent],
                    hoverinfo="text",
                ))
            fig.update_layout(
                plot_bgcolor=BG, paper_bgcolor=BG,
                font=dict(color=TEXT, family="Arial", size=11),
                legend=dict(bgcolor=PANEL, bordercolor=BORDER, borderwidth=1),
                margin=dict(l=8, r=8, t=8, b=8),
                xaxis=dict(gridcolor=PANEL2, zerolinecolor=PANEL2),
                yaxis=dict(gridcolor=PANEL2, zerolinecolor=PANEL2,
                           autorange="reversed"),
            )
            st.plotly_chart(fig, use_container_width=True)

    # ---------------------------------------------------------- Tab 4
    with t4:
        st.subheader("Só as arestas que cruzam projeto")
        inter = [e for e in edges if nos[e["o"]]["projeto"] != nos[e["d"]]["projeto"]]
        st.caption(f"{len(inter)} de {len(edges)} arestas são inter-projeto.")
        # subset de nós envolvidos
        ids_inter = {e["o"] for e in inter} | {e["d"] for e in inter}
        nos_inter = {d: nos[d] for d in ids_inter}
        st.plotly_chart(
            _timeline_conectada(nos_inter, inter),
            use_container_width=True, config={"displayModeBar": False},
        )
        with st.expander("Listar arestas inter-projeto", expanded=False):
            for e in inter:
                st.markdown(f"- **{nos[e['o']]['atividade'][:44]}** "
                            f"({nos[e['o']]['projeto']}) → "
                            f"**{nos[e['d']]['atividade'][:44]}** "
                            f"({nos[e['d']]['projeto']}) — "
                            f"*{e['tipo']}/{e['forca']}* · {e['nota']}")
        st.divider()
        st.subheader("🚩 Flags — pendências")
        if flags:
            for f in flags:
                st.markdown(f"- **{f['item']}** *({f['tipo']})* — {f['nota']}")
        else:
            st.caption("_sem flags_")


if __name__ == "__main__":
    main()
