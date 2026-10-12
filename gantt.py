"""Deriva Alocacao_Gantt (pessoa x projeto x span de meses) das Atividades.

Regras (engenharia reversa validada contra o gantt de referência, 125/125):
- taxonomia de gantt mais granular que 'projeto' (PRiME por frente, TCE por
  assessoria/curso, Bens Públicos separando MiniGuia).
- meses de uma atividade = mês do `prazo` (ano < 2026 -> 1, ano > 2026 -> 12)
  + todos os meses citados em `prazo_obs`.
- cobertura POR PROJETO: se < 30% das linhas (pessoa x atividade) tiverem mês,
  todas as linhas do projeto viram 1..12 [sem_data_no_cronograma].
- senão, por pessoa: min..max dos seus meses [cronograma]; se a pessoa não tem
  nenhum mês, herda o span do projeto [cronograma_sem_prazo].
- linhas de projetos sem cronograma preservadas do gantt antigo [alocacao_equipe]
"""
import re
import pandas as pd

COV_MIN = 0.30

MESES = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}


def meses_do_texto(txt):
    if not isinstance(txt, str):
        return []
    t = txt.lower()
    return [m for pref, m in MESES.items() if re.search(r"\b" + pref, t)]


def prime_frente(sub):
    s = str(sub or "")
    if re.search(r"M1 \(Fundamentals\).*M2|M1.*M2.*M3", s):
        return "PRiME — Multi-módulo"
    if "Fundamentals" in s:
        return "PRiME — Fundamentals"
    if "Hands-on" in s:
        return "PRiME — Hands-on"
    if "M3 (IE)" in s or s.startswith("Módulo"):
        return "PRiME — Av. Impacto (IE)"
    if "Innovation WS" in s:
        return "PRiME — Oficinas Inovação"
    if "C3" in s or "Comms" in s:
        return "PRiME — Comms/KM"
    return "PRiME — Gestão"


def gantt_projeto(row):
    p, sub = row["projeto"], str(row["sub_projeto"] or "")
    if p == "PRiME III":
        return prime_frente(sub)
    if p == "TCE ES":
        return "TCE ES — Curso" if "Curso" in sub else "TCE ES — Assessoria"
    if p == "Bens Públicos":
        return "Bens Públicos — MiniGuia" if sub == "MiniGuia" else "Bens Públicos"
    return p


def derivar(ativ, gantt_antigo, canon):
    a = ativ.copy().reset_index(drop=True)
    a["proj_g"] = a.apply(gantt_projeto, axis=1)
    pz = pd.to_datetime(a["prazo"], errors="coerce")

    meses_col = []
    for i in range(len(a)):
        ms = []
        d = pz.iloc[i]
        if pd.notna(d):
            ms.append(1 if d.year < 2026 else (12 if d.year > 2026 else d.month))
        ms += meses_do_texto(a["prazo_obs"].iloc[i])
        meses_col.append(sorted(set(ms)))
    a["meses"] = meses_col

    linhas = []
    for _, r in a.iterrows():
        for pes in [x.strip() for x in str(r["responsaveis"]).split(";") if x.strip()]:
            if pes in canon:
                linhas.append(dict(pessoa=pes, projeto=r["proj_g"], meses=r["meses"]))
    df = pd.DataFrame(linhas)

    out = []
    for proj, gp in df.groupby("projeto"):
        todos = [m for ms in gp["meses"] for m in ms]
        cov = sum(1 for ms in gp["meses"] if ms) / len(gp)
        baixa = cov < COV_MIN or not todos
        pspan = (min(todos), max(todos)) if todos else (1, 12)
        for pes, g in gp.groupby("pessoa"):
            n = len(g)
            ms = [m for x in g["meses"] for m in x]
            if baixa:
                de, ate, origem = 1, 12, "sem_data_no_cronograma"
            elif ms:
                de, ate, origem = min(ms), max(ms), "cronograma"
            else:
                de, ate, origem = pspan[0], pspan[1], "cronograma_sem_prazo"
            out.append(dict(pessoa=pes, projeto=proj, mes_de=int(de), mes_ate=int(ate),
                            origem=origem, n_atividades=n))

    novo = pd.DataFrame(out)
    manual = gantt_antigo[gantt_antigo["origem"] == "alocacao_equipe"].copy()
    res = pd.concat([novo, manual], ignore_index=True)
    return res.sort_values(["pessoa", "projeto"]).reset_index(drop=True)
