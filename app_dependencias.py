"""
CLEAR 2026 — Cadeia de Dependências
App Streamlit SEPARADO (segundo main file no mesmo repo caiodesouzacastro/clear-2026).
Deploy: share.streamlit.io -> New app -> repo clear-2026 -> Main file: app_dependencias.py
       -> ganha URL própria. Usa a mesma senha (st.secrets["app_password"]).

Lê, do próprio repo:
  - CLEAR_Master_2026.xlsx        (status/prazo AO VIVO por (projeto, atividade, detalhe))
  - Dependencias_CLEAR_2026.xlsx  (abas Nos + Arestas)

Convenção: aresta origem -> destino = "origem DESTRAVA destino".
"""

import re
import collections
from pathlib import Path

import streamlit as st
import openpyxl

BASE = Path(__file__).parent
MASTER = BASE / "CLEAR_Master_2026.xlsx"
DEPS = BASE / "Dependencias_CLEAR_2026.xlsx"

# projetos com fases numeradas realmente sequenciais (medido) -> derivam "provável"
DERIVAVEIS = {"Painel CLEAR", "EVALAC"}
NUMPAT = re.compile(r"^\s*(?:fase\s*)?[sc]?\s*(\d+)\b", re.I)

STATUS_COR = {
    "concluído": "#2E7D32", "pronto": "#2E7D32", "enviado para eesp": "#2E7D32",
    "em andamento": "#1565C0",
    "atrasado": "#C62828",
    "não iniciado": "#9E9E9E", "reunião": "#9E9E9E",
}
COR_FADE = "#E8E8E8"
COR_PLACEHOLDER = "#FCE4D6"


# ----------------------------------------------------------------------------- auth
def _senha_ok() -> bool:
    if st.session_state.get("_auth"):
        return True
    pw = st.secrets.get("app_password") if hasattr(st, "secrets") else None
    if not pw:  # sem secret configurado -> não bloqueia (dev)
        return True
    val = st.text_input("Senha", type="password")
    if val == pw:
        st.session_state["_auth"] = True
        return True
    if val:
        st.error("Senha incorreta.")
    return False


# ----------------------------------------------------------------------------- dados
@st.cache_data(show_spinner=False)
def carregar():
    # --- Master: chave natural -> {status, prazo} atuais ---
    wm = openpyxl.load_workbook(MASTER, read_only=True, data_only=True)
    ws = wm["Atividades"]
    hdr = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    H = {h: i for i, h in enumerate(hdr)}
    master = {}
    master_por_proj = collections.defaultdict(list)
    for r in ws.iter_rows(min_row=2, values_only=True):
        k = (r[H["projeto"]], r[H["atividade"]], r[H["detalhe"]])
        info = {
            "status": r[H["status"]],
            "prazo": r[H["prazo"]],
            "sub": r[H["sub_projeto"]],
            "eh_ent": r[H["eh_entregavel"]],
        }
        master[k] = info
        master_por_proj[r[H["projeto"]]].append((r[H["atividade"]], info))

    # --- Dependências: Nos + Arestas ---
    wd = openpyxl.load_workbook(DEPS, read_only=True, data_only=True)
    nos = {}
    for r in wd["Nos"].iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        did = r[0]
        proj, ativ, det, prazo_dep, fonte = r[1], r[2], r[3], r[4], r[6]
        m = master.get((proj, ativ, det))
        nos[did] = {
            "projeto": proj, "atividade": ativ, "detalhe": det,
            "fonte": fonte,
            "orfao": (fonte != "placeholder") and (m is None),
            "status": (m["status"] if m else None),
            "prazo": (m["prazo"] if m else prazo_dep),
        }
    edges = []
    for r in wd["Arestas"].iter_rows(min_row=2, values_only=True):
        if not r[0] or not r[1]:
            continue
        edges.append({"o": r[0], "d": r[1], "tipo": r[2] or "insumo",
                      "forca": r[3] or "dura", "nota": r[5] or ""})

    # --- camada "provável" derivada: fases numeradas de Painel/EVALAC ---
    # entre os NÓS (entregáveis) desses projetos, liga na ordem de prazo dentro do projeto.
    derivadas = []
    for proj in DERIVAVEIS:
        ns = [d for d, n in nos.items() if n["projeto"] == proj and not n["orfao"]]
        ns.sort(key=lambda d: (nos[d]["prazo"] is None, nos[d]["prazo"] or ""))
        for a, b in zip(ns, ns[1:]):
            derivadas.append({"o": a, "d": b})

    return nos, edges, derivadas


def _cor_status(n):
    if n["fonte"] == "placeholder":
        return COR_PLACEHOLDER
    s = (n["status"] or "").strip().lower()
    return STATUS_COR.get(s, "#BDBDBD")


def _label(did, n):
    ativ = (n["atividade"] or "")[:34]
    prazo = n["prazo"]
    if hasattr(prazo, "date"):
        prazo = prazo.date().isoformat()
    prazo = prazo or "s/ data"
    st_txt = "PLACEHOLDER" if n["fonte"] == "placeholder" else (n["status"] or "?")
    ativ = ativ.replace('"', "'").replace("\n", " ")
    return f"{did}\\n{ativ}\\n[{st_txt} · {prazo}]"


# ---------------------------------------------------------------- grafo / DOT
def _grafo(edges):
    adj = collections.defaultdict(list)
    radj = collections.defaultdict(list)
    for e in edges:
        adj[e["o"]].append(e["d"])
        radj[e["d"]].append(e["o"])
    return adj, radj


def _downstream(adj, s):
    seen, stack = set(), [s]
    while stack:
        u = stack.pop()
        for v in adj[u]:
            if v not in seen:
                seen.add(v)
                stack.append(v)
    return seen


def _upstream(radj, s):
    seen, stack = set(), [s]
    while stack:
        u = stack.pop()
        for v in radj[u]:
            if v not in seen:
                seen.add(v)
                stack.append(v)
    return seen


def _caminho_mais_longo(radj, t):
    memo = {}

    def dfs(u):
        if u in memo:
            return memo[u]
        if not radj[u]:
            memo[u] = [u]
            return memo[u]
        best = max((dfs(p) for p in radj[u]), key=len)
        memo[u] = best + [u]
        return memo[u]

    return dfs(t)


def _dot(nos, edges, derivadas, *, foco=None, destaque_nos=None,
         destaque_arestas=None, mostrar_derivadas=False, clusters=False):
    destaque_nos = destaque_nos or set()
    destaque_arestas = destaque_arestas or set()
    tem_destaque = bool(destaque_nos)

    linhas = ['digraph G {', 'rankdir=LR;', 'bgcolor="transparent";',
              'node [shape=box style="filled,rounded" fontname="Arial" fontsize=10 '
              'penwidth=1 color="#B0B0B0"];',
              'edge [fontname="Arial" fontsize=8 color="#888888"];']

    # nós que aparecem
    usados = set()
    for e in edges:
        usados.add(e["o"]); usados.add(e["d"])
    if mostrar_derivadas:
        for e in derivadas:
            usados.add(e["o"]); usados.add(e["d"])

    def node_line(did):
        n = nos[did]
        cor = _cor_status(n)
        fontcor = "#FFFFFF" if cor in ("#2E7D32", "#1565C0", "#C62828") else "#000000"
        pen, bordercor = 1, "#B0B0B0"
        if tem_destaque and did not in destaque_nos:
            cor, fontcor, bordercor = COR_FADE, "#9E9E9E", "#D5D5D5"
        if did == foco:
            pen, bordercor = 3, "#000000"
        elif did in destaque_nos:
            pen, bordercor = 2, "#333333"
        extra = ' style="filled,rounded,dashed"' if n["fonte"] == "placeholder" else ""
        return (f'"{did}" [label="{_label(did, n)}" fillcolor="{cor}" '
                f'fontcolor="{fontcor}" color="{bordercor}" penwidth={pen}{extra}];')

    if clusters:
        porproj = collections.defaultdict(list)
        for did in usados:
            porproj[nos[did]["projeto"]].append(did)
        for i, (proj, dids) in enumerate(sorted(porproj.items())):
            linhas.append(f'subgraph cluster_{i} {{ label="{proj}"; '
                          f'style="rounded"; color="#CCCCCC"; fontname="Arial"; fontsize=11;')
            for did in dids:
                linhas.append(node_line(did))
            linhas.append("}")
    else:
        for did in usados:
            linhas.append(node_line(did))

    # arestas declaradas
    for e in edges:
        estilo = "solid" if e["forca"] == "dura" else "dashed"
        cor = "#555555" if e["tipo"] == "bloqueio" else "#999999"
        pen = 1
        if (e["o"], e["d"]) in destaque_arestas:
            cor, pen = "#C62828", 2
        elif tem_destaque:
            cor = "#DDDDDD"
        linhas.append(f'"{e["o"]}" -> "{e["d"]}" [style={estilo} color="{cor}" penwidth={pen}];')

    # arestas prováveis (derivadas)
    if mostrar_derivadas:
        for e in derivadas:
            linhas.append(f'"{e["o"]}" -> "{e["d"]}" '
                          f'[style=dotted color="#C0A0C0" penwidth=1 constraint=false];')

    linhas.append("}")
    return "\n".join(linhas)


# ----------------------------------------------------------------------------- UI
def main():
    st.set_page_config(page_title="CLEAR — Cadeia de Dependências", layout="wide")
    st.title("CLEAR 2026 — Cadeia de Dependências")
    st.caption("Aresta A → B lê-se: **A destrava B** (B depende de A). "
               "Cor do nó = status atual no Master. Linha cheia = dura · tracejada = mole · "
               "pontilhada roxa = precedência provável (derivada).")

    if not _senha_ok():
        st.stop()

    try:
        nos, edges, derivadas = carregar()
    except FileNotFoundError as e:
        st.error(f"Arquivo não encontrado no repo: {e.filename}. "
                 "Confirme que CLEAR_Master_2026.xlsx e Dependencias_CLEAR_2026.xlsx "
                 "estão na raiz do repositório.")
        st.stop()

    orfaos = [d for d, n in nos.items() if n["orfao"]]
    if orfaos:
        st.warning("⚠️ Âncoras órfãs (dep_id que não resolve no Master — provável rename da "
                   f"atividade): {', '.join(orfaos)}. Corrija o texto na aba Nos ou no Master.")

    adj, radj = _grafo(edges)
    with st.sidebar:
        st.header("Opções")
        vista = st.radio("Tela", ["Impacto a jusante",
                                  "Caminho crítico por linha de chegada",
                                  "Mapa inter-projetos"])
        mostrar_deriv = st.checkbox("Mostrar precedência provável (Painel/EVALAC)", value=False)
        st.divider()
        st.caption(f"{len(nos)} nós · {len(edges)} arestas declaradas · "
                   f"{len(orfaos)} órfãs")

    def rotulo(d):
        return f"{d} — {(nos[d]['atividade'] or '')[:40]}"

    # -------------------------------------------------- Tela 1
    if vista == "Impacto a jusante":
        st.subheader("Se isto atrasar, o que cai junto?")
        candidatos = [d for d in nos if adj[d] or radj[d]]
        candidatos.sort()
        alvo = st.selectbox("Atividade", candidatos, format_func=rotulo)
        down = _downstream(adj, alvo)
        destaque = {alvo} | down
        arestas_dest = {(e["o"], e["d"]) for e in edges
                        if e["o"] in destaque and e["d"] in destaque}
        chegadas = [d for d in down if not adj[d]]
        c1, c2 = st.columns([3, 1])
        with c2:
            st.metric("Atividades afetadas", len(down))
            st.markdown("**Linhas de chegada atingidas:**")
            st.markdown("\n".join(f"- {d}" for d in sorted(chegadas)) or "_nenhuma_")
        with c1:
            st.graphviz_chart(_dot(nos, edges, derivadas, foco=alvo,
                                   destaque_nos=destaque, destaque_arestas=arestas_dest,
                                   mostrar_derivadas=mostrar_deriv), use_container_width=True)

    # -------------------------------------------------- Tela 2
    elif vista == "Caminho crítico por linha de chegada":
        st.subheader("A corrente mais longa que sustenta cada marco")
        chegadas = sorted(d for d in nos if not adj[d] and radj[d])
        alvo = st.selectbox("Linha de chegada", chegadas, format_func=rotulo)
        caminho = _caminho_mais_longo(radj, alvo)
        destaque = set(caminho)
        arestas_dest = set(zip(caminho, caminho[1:]))
        st.markdown("**Caminho:** " + "  →  ".join(caminho))
        st.graphviz_chart(_dot(nos, edges, derivadas, foco=alvo,
                               destaque_nos=destaque, destaque_arestas=arestas_dest,
                               mostrar_derivadas=mostrar_deriv), use_container_width=True)

    # -------------------------------------------------- Tela 3
    else:
        st.subheader("Só as arestas que cruzam projeto")
        inter = [e for e in edges if nos[e["o"]]["projeto"] != nos[e["d"]]["projeto"]]
        st.caption(f"{len(inter)} de {len(edges)} arestas são inter-projeto — "
                   "são as mais caras quando quebram.")
        st.graphviz_chart(_dot(nos, inter, derivadas, clusters=True,
                               mostrar_derivadas=False), use_container_width=True)
        with st.expander("Listar arestas inter-projeto"):
            for e in inter:
                st.markdown(f"- `{e['o']}` ({nos[e['o']]['projeto']}) → "
                            f"`{e['d']}` ({nos[e['d']]['projeto']}) — "
                            f"*{e['tipo']}/{e['forca']}* · {e['nota']}")


if __name__ == "__main__":
    main()
