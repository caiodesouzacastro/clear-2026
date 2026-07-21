"""
CLEAR 2026 — Cadeia de Dependências (v2)
Tema escuro, paleta igual ao app principal (config.toml do repo).
Quatro abas:
  1. Impacto a jusante  (grafo enxuto: 63 entregáveis, arestas declaradas)
  2. Caminho crítico    (mesmo grafo, filtro por linha de chegada)
  3. Todas as atividades (timeline Gantt: as 736 do Master, por projeto)
  4. Mapa inter-projetos + Flags
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

# ---- paleta idêntica ao app principal (config.toml + CORES_STATUS de app.py)
BG        = "#0A1929"
PANEL     = "#132F4C"
PANEL2    = "#1E3A5F"
PRIMARY   = "#5090D3"
TEXT      = "#FFFFFF"
MUTED     = "#B2BAC2"
BORDER    = "#26456B"

STATUS_COR = {
    "concluído":         "#6F7E8C",
    "pronto":            "#6F7E8C",
    "enviado para eesp": "#6F7E8C",
    "em andamento":      "#5090D3",
    "não iniciado":      "#3A5169",
    "atrasado":          "#7BB3F0",
    "reunião":           "#4A6FA5",
}
COR_PLACEHOLDER = "#4A6FA5"
COR_DESCONHECIDO = "#3A5169"
COR_FADE = "#1E3A5F"

DERIVAVEIS = {"Painel CLEAR", "EVALAC"}


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
        info = {
            "status": r[H["status"]],
            "prazo": r[H["prazo"]],
            "sub": r[H["sub_projeto"]],
            "eh_ent": r[H["eh_entregavel"]] is True,
            "resp": r[H["responsaveis"]],
        }
        master[(proj, ativ, det)] = info
        atividades.append({
            "projeto": proj or "(sem projeto)",
            "sub": info["sub"] or "",
            "atividade": ativ or "",
            "detalhe": det or "",
            "prazo": info["prazo"],
            "status": info["status"] or "Não Iniciado",
            "responsaveis": info["resp"] or "",
            "eh_entregavel": info["eh_ent"],
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
                      "tipo": r[2] or "insumo",
                      "forca": r[3] or "dura",
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


def _cor_status(status_txt, fonte=None):
    if fonte == "placeholder":
        return COR_PLACEHOLDER
    if not status_txt:
        return COR_DESCONHECIDO
    return STATUS_COR.get(status_txt.strip().lower(), COR_DESCONHECIDO)


# ============================================================ grafo
def _grafo(edges):
    adj = collections.defaultdict(list)
    radj = collections.defaultdict(list)
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


def _dot(nos, edges, derivadas, *, foco=None, destaque_nos=None,
         destaque_arestas=None, mostrar_derivadas=False, clusters=False):
    destaque_nos = destaque_nos or set()
    destaque_arestas = destaque_arestas or set()
    tem_destaque = bool(destaque_nos)

    linhas = [
        'digraph G {',
        'rankdir=LR;',
        f'bgcolor="{BG}";',
        f'node [shape=box style="filled,rounded" fontname="Arial" fontsize=10 '
        f'penwidth=1 color="{BORDER}"];',
        f'edge [fontname="Arial" fontsize=8 color="{MUTED}"];',
    ]
    usados = set()
    for e in edges: usados |= {e["o"], e["d"]}
    if mostrar_derivadas:
        for e in derivadas: usados |= {e["o"], e["d"]}

    def node_line(did):
        n = nos[did]
        cor = _cor_status(n["status"], n["fonte"])
        ativ = (n["atividade"] or "")[:26].replace('"', "'").replace("\n", " ")
        prazo = n["prazo"]
        if hasattr(prazo, "date"): prazo = prazo.date().isoformat()
        prazo = prazo or "s/ data"
        st_txt = "PLACEHOLDER" if n["fonte"] == "placeholder" else (n["status"] or "?")
        tooltip = f"{did} | {n['projeto']} | {ativ} | [{st_txt} · {prazo}]"
        label = f"{did}\\n{ativ}"
        fontcor = TEXT
        pen, bordercor = 1, BORDER
        if tem_destaque and did not in destaque_nos:
            cor, fontcor, bordercor = COR_FADE, MUTED, BORDER
        if did == foco:
            pen, bordercor = 3, PRIMARY
        elif did in destaque_nos:
            pen, bordercor = 2, PRIMARY
        extra = ' style="filled,rounded,dashed"' if n["fonte"] == "placeholder" else ""
        return (f'"{did}" [label="{label}" tooltip="{tooltip}" fillcolor="{cor}" '
                f'fontcolor="{fontcor}" color="{bordercor}" penwidth={pen}{extra}];')

    if clusters:
        porproj = collections.defaultdict(list)
        for did in usados: porproj[nos[did]["projeto"]].append(did)
        for i, (proj, dids) in enumerate(sorted(porproj.items())):
            linhas.append(
                f'subgraph cluster_{i} {{ label="{proj}"; style="rounded"; '
                f'color="{PRIMARY}"; fontcolor="{TEXT}"; fontname="Arial"; fontsize=11;'
            )
            for did in dids: linhas.append(node_line(did))
            linhas.append("}")
    else:
        for did in usados: linhas.append(node_line(did))

    for e in edges:
        estilo = "solid" if e["forca"] == "dura" else "dashed"
        cor = PRIMARY if e["tipo"] == "bloqueio" else MUTED
        pen = 1
        if (e["o"], e["d"]) in destaque_arestas:
            cor, pen = PRIMARY, 2.5
        elif tem_destaque:
            cor = BORDER
        linhas.append(f'"{e["o"]}" -> "{e["d"]}" [style={estilo} color="{cor}" penwidth={pen}];')

    if mostrar_derivadas:
        for e in derivadas:
            linhas.append(f'"{e["o"]}" -> "{e["d"]}" '
                          f'[style=dotted color="{MUTED}" penwidth=1 constraint=false];')
    linhas.append("}")
    return "\n".join(linhas)


# ============================================================ UI
def main():
    st.set_page_config(page_title="CLEAR 2026 · Dependências",
                       page_icon="🔗", layout="wide",
                       initial_sidebar_state="collapsed")

    st.markdown(f"""
    <style>
      .block-container {{ padding-top: 1.5rem; padding-bottom: 1rem; max-width: 100%; }}
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
    st.caption("Aresta A → B lê-se: **A destrava B**. "
               "Cor do nó = status atual no Master · "
               "linha cheia = dura · tracejada = mole · pontilhada = provável (derivada).")

    if not _senha_ok():
        st.stop()

    try:
        nos, edges, derivadas, atividades, flags = carregar()
    except FileNotFoundError as e:
        st.error(f"Arquivo não encontrado: {e.filename}")
        st.stop()

    orfaos = [d for d, n in nos.items() if n["orfao"]]
    if orfaos:
        st.warning(f"⚠️ Nós órfãos (rename provável): {', '.join(orfaos)}")

    adj, radj = _grafo(edges)

    def rotulo(d):
        return f"{d} — {(nos[d]['atividade'] or '')[:44]}"

    t1, t2, t3, t4 = st.tabs([
        "⬇️ Impacto a jusante",
        "🎯 Caminho crítico",
        "📅 Todas as atividades",
        "🔀 Inter-projetos & flags",
    ])

    # ---------------------------------------------------------- Tab 1
    with t1:
        col_ctrl, col_info = st.columns([3, 1])
        candidatos = sorted(d for d in nos if adj[d] or radj[d])
        with col_ctrl:
            alvo = st.selectbox("Escolha uma atividade:", candidatos,
                                format_func=rotulo, key="t1_sel")
            mostrar_deriv = st.checkbox(
                "Mostrar precedência provável (Painel/EVALAC)", key="t1_der")
        down = _downstream(adj, alvo)
        destaque = {alvo} | down
        arestas_dest = {(e["o"], e["d"]) for e in edges
                        if e["o"] in destaque and e["d"] in destaque}
        chegadas = sorted(d for d in down if not adj[d])
        with col_info:
            st.metric("Atividades afetadas", len(down))
            st.markdown("**Linhas de chegada atingidas:**")
            if chegadas:
                for d in chegadas:
                    st.markdown(f"- `{d}` — {nos[d]['atividade'][:34]}")
            else:
                st.caption("_nenhuma_")
        st.graphviz_chart(
            _dot(nos, edges, derivadas, foco=alvo,
                 destaque_nos=destaque, destaque_arestas=arestas_dest,
                 mostrar_derivadas=mostrar_deriv),
            use_container_width=True,
        )

    # ---------------------------------------------------------- Tab 2
    with t2:
        chegadas = sorted(d for d in nos if not adj[d] and radj[d])
        alvo = st.selectbox("Linha de chegada:", chegadas,
                            format_func=rotulo, key="t2_sel")
        caminho = _caminho_mais_longo(radj, alvo)
        destaque = set(caminho)
        arestas_dest = set(zip(caminho, caminho[1:]))
        st.markdown("**Caminho crítico:** " + " → ".join(f"`{c}`" for c in caminho))
        st.graphviz_chart(
            _dot(nos, edges, derivadas, foco=alvo,
                 destaque_nos=destaque, destaque_arestas=arestas_dest),
            use_container_width=True,
        )

    # ---------------------------------------------------------- Tab 3 — todas
    with t3:
        st.markdown("**Timeline de todas as atividades do Master.** "
                    "Cada barra = uma atividade, colorida por status. "
                    "Diamantes ♦ marcam entregáveis.")
        projs = sorted({a["projeto"] for a in atividades if a["projeto"]})
        f1, f2, f3 = st.columns([2, 2, 1])
        with f1:
            sel_projs = st.multiselect("Projeto(s)", projs, default=projs)
        with f2:
            sel_status = st.multiselect(
                "Status",
                ["Concluído", "Em Andamento", "Não Iniciado", "Atrasado", "Reunião"],
                default=["Em Andamento", "Não Iniciado", "Atrasado"],
            )
        with f3:
            so_ent = st.checkbox("Só entregáveis", value=False)

        filt = [a for a in atividades
                if a["projeto"] in sel_projs
                and a["prazo"] is not None
                and str(a["status"]).strip() in sel_status
                and (not so_ent or a["eh_entregavel"])]
        sem_data = sum(1 for a in atividades
                       if a["projeto"] in sel_projs and a["prazo"] is None
                       and (not so_ent or a["eh_entregavel"]))
        st.caption(f"{len(filt)} atividades no gráfico · {sem_data} sem data (ocultas)")

        if not filt:
            st.info("Nada a mostrar com esse filtro.")
        else:
            for a in filt:
                a["_start"] = a["prazo"]
                a["_end"] = a["prazo"] + dt.timedelta(days=1)
                a["_lbl"] = f"{a['atividade'][:60]}"
                a["_status_norm"] = str(a["status"]).strip()
            fig = px.timeline(
                filt, x_start="_start", x_end="_end", y="projeto",
                color="_status_norm",
                color_discrete_map={
                    "Concluído": "#6F7E8C",
                    "Em Andamento": "#5090D3",
                    "Não Iniciado": "#3A5169",
                    "Atrasado": "#7BB3F0",
                    "Reunião": "#4A6FA5",
                },
                hover_data={"_lbl": True, "responsaveis": True, "sub": True,
                            "_start": False, "_end": False, "projeto": False,
                            "_status_norm": True},
                height=max(360, 34 * len(sel_projs) + 120),
            )
            ent = [a for a in filt if a["eh_entregavel"]]
            if ent:
                fig.add_trace(go.Scatter(
                    x=[a["prazo"] for a in ent],
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
                yaxis=dict(gridcolor=PANEL2, zerolinecolor=PANEL2, autorange="reversed"),
            )
            st.plotly_chart(fig, use_container_width=True)

    # ---------------------------------------------------------- Tab 4
    with t4:
        st.subheader("Só as arestas que cruzam projeto")
        inter = [e for e in edges if nos[e["o"]]["projeto"] != nos[e["d"]]["projeto"]]
        st.caption(f"{len(inter)} de {len(edges)} arestas são inter-projeto.")
        st.graphviz_chart(_dot(nos, inter, derivadas, clusters=True),
                          use_container_width=True)
        with st.expander("Listar arestas inter-projeto", expanded=False):
            for e in inter:
                st.markdown(f"- `{e['o']}` ({nos[e['o']]['projeto']}) → "
                            f"`{e['d']}` ({nos[e['d']]['projeto']}) — "
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
