"""Parser robusto para abas no formato padrão de cronograma CLEAR.
Extrai linhas no schema do master: sub_projeto, atividade, detalhe,
responsaveis(_raw), n_responsaveis, prazo, prazo_obs, status, eh_entregavel,
presencial.
"""
import re
import datetime as dt
import openpyxl

# --- Normalização de nomes (regras duras do CLEAR) ---
TEAM = ["Bia B","Bia S","Caio","Carol","Cecilia","Fabrícia","Fred","Hisrael",
        "Julia","Junior","Lorena","Luan","Luigi","Lycia","Michel","Samu"]

# mapeia variantes -> canônico
NAME_MAP = {
    "zinho":"Michel", "pleno":"Fred", "senior":"Hisrael", "sênior":"Hisrael",
    "luiggi":"Luigi", "luigui":"Luigi",
    "fabricia":"Fabrícia", "fabrícia":"Fabrícia",
    "cecília":"Cecilia", "cecilia":"Cecilia",
    "júlia":"Julia", "julia":"Julia",
    "júnior":"Junior", "junior":"Junior",
    "bia":"Bia",  # ambíguo: tratado à parte abaixo
    "bia b":"Bia B", "bia s":"Bia S",
    "caraol":"Carol", "carol":"Carol",
    "lycia":"Lycia","lorena":"Lorena","luan":"Luan","michel":"Michel",
    "samu":"Samu","caio":"Caio","fred":"Fred","hisrael":"Hisrael","luigi":"Luigi",
}
# entidades externas mantidas como estão (não são pessoas do time)
EXTERNAL = {"ifad","ove","ove/bid","bid","tce","sicredi","comitê gestor","comitê",
            "comite gestor","afrea","onu mulheres","cmf-bid","alide","ti fgv",
            "biblioteca","equipe implementadora","clear lac","clear lab","ove-ecd",
            "comunicação fgv","paulinha","hugo","isabella","bárbara","barbara"}
DISCARD = {"pmo"}

MESES = {"janeiro":1,"fevereiro":2,"março":3,"marco":3,"abril":4,"maio":5,"junho":6,
         "julho":7,"agosto":8,"setembro":9,"outubro":10,"novembro":11,"dezembro":12,
         "jan":1,"fev":2,"mar":3,"abr":4,"mai":5,"jun":6,"jul":7,"ago":8,"set":9,
         "out":10,"nov":11,"dez":12}

STATUS_SET = {"concluído","concluido","em andamento","não iniciado","nao iniciado",
              "entregável","entregavel","reunião","reuniao","em produção","publicada",
              "publicado","em finalização","em revisão"}


def norm_one(raw):
    """Normaliza um único nome para o canônico do time, ou retorna title-case."""
    if raw is None:
        return None
    s = str(raw).strip().strip(".").strip("?").strip()
    s = re.sub(r"\s+", " ", s)
    if not s:
        return None
    low = s.lower()
    if low in DISCARD:
        return None
    if low in NAME_MAP:
        return NAME_MAP[low]
    # nomes externos -> preserva capitalização original enxuta
    if low in EXTERNAL:
        return s
    # casos "Bia" isolado sem sufixo -> mantém "Bia" (ambíguo; não inventar)
    if low == "bia":
        return "Bia"
    # title-case padrão para qualquer outro
    return s


CANON = {"bia b":"Bia B","bia s":"Bia S","caio":"Caio","carol":"Carol",
         "cecilia":"Cecilia","cecília":"Cecilia","fabricia":"Fabrícia","fabrícia":"Fabrícia",
         "fred":"Fred","hisrael":"Hisrael","julia":"Julia","júlia":"Julia","ju":"Julia",
         "junior":"Junior","júnior":"Junior","lorena":"Lorena","luan":"Luan",
         "luigi":"Luigi","luiggi":"Luigi","lycia":"Lycia","michel":"Michel","zinho":"Michel",
         "samu":"Samu","samy":"Samu","pleno":"Fred","senior":"Hisrael","sênior":"Hisrael",
         "beatriz b":"Bia B","bai b":"Bia B","caraol":"Carol",
         "bias":"Bia S","biab":"Bia B","bia. s":"Bia S","bia. b":"Bia B"}
# externos reais que devem ser preservados como responsáveis
EXTERNAL_KEEP = {"marcio":"Marcio","lara":"Lara","lia":"Lia","amanda":"Amanda","alei":"Alei",
                 "hugo":"Hugo","gus":"Gustavo","gustavo":"Gustavo","paulinha":"Paulinha",
                 "barbara":"Bárbara","bárbara":"Bárbara","livia":"Livia","tania":"Tania",
                 "cristina navas":"Cristina Navas","gabriel w":"Gabriel W","ifad":"IFAD",
                 "ove":"OVE","bid":"BID","ove/bid":"OVE/BID","sicredi":"Sicredi","tce":"TCE",
                 "comitê gestor":"Comitê Gestor","comite gestor":"Comitê Gestor",
                 "comitê":"Comitê Gestor","aulasneo":"Aulasneo","cross content":"Cross Content",
                 "crosscontent":"Cross Content","clear":"CLEAR","clear lac":"CLEAR LAC",
                 "clear lab":"CLEAR LAB","rede clear":"Rede CLEAR","onu mulheres":"ONU Mulheres",
                 "afrea":"AfrEA","ti fgv":"TI FGV","biblioteca":"Biblioteca",
                 "equipe implementadora":"Equipe Implementadora","cmf-bid":"CMF-BID",
                 "amanda":"Amanda","alei":"Alei","auditores":"Auditores","hugo":"Hugo",
                 "comunicação fgv":"Comunicação FGV","comunicacao fgv":"Comunicação FGV",
                 "isabella":"Isabella","ana":"Ana"}
_NOTE_DROP = {"dummy","confirmar","apoio","equipe","provavelmente nós","bia"}


def split_responsaveis(raw):
    """Limpa e separa responsáveis. Remove parênteses/notas, mapeia variantes,
    mantém time canônico + externos reais, descarta fragmentos."""
    if raw is None:
        return [], ""
    raw_s = str(raw).strip()
    s = re.sub(r"\([^)]*\)", " ", raw_s)          # remove parênteses balanceados
    s = re.sub(r"\([^)]*$", " ", s)               # remove parêntese aberto pendente
    s = re.sub(r"\b(com (ajuda|apoio)[^,;/+]*|supervisionad[ao] por[^,;/+]*|sup\.[^,;/+]*|cc [^,;/+]*|obs:[^,;/+]*)",
               " ", s, flags=re.I)
    parts = re.split(r"[;/+,&\n]| e ", s)
    out = []
    for p in parts:
        t = re.sub(r"\s+", " ", p).strip(" .)?(").strip()
        if not t:
            continue
        low = t.lower()
        if low in _NOTE_DROP or low in DISCARD:
            continue
        if ":" in t or any(ch.isdigit() for ch in t) or len(t) > 26:
            continue
        if low in CANON:
            name = CANON[low]
        elif low in EXTERNAL_KEEP:
            name = EXTERNAL_KEEP[low]
        else:
            continue  # token desconhecido/fragmento -> descarta
        if name and name not in out:
            out.append(name)
    return out, raw_s


def parse_prazo(val):
    """Retorna (date|None, obs|None). Datas viram datetime; textos viram obs."""
    if val is None:
        return None, None
    if isinstance(val, (dt.datetime, dt.date)):
        d = val if isinstance(val, dt.datetime) else dt.datetime(val.year, val.month, val.day)
        return d, None
    s = str(val).strip().replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    if not s:
        return None, None
    # dd/mm/yyyy ou dd/mm/yy
    m = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b", s)
    if m:
        dd, mm, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if yy < 100:
            yy += 2000
        try:
            return dt.datetime(yy, mm, dd), (s if len(s) > 12 else None)
        except ValueError:
            pass
    # "22 de junho" / "11 de setembro" -> usa ano 2026
    m = re.search(r"\b(\d{1,2})\s*(?:de\s+)?([a-zç]+)", s.lower())
    if m and m.group(2) in MESES:
        dd = int(m.group(1)); mm = MESES[m.group(2)]
        try:
            return dt.datetime(2026, mm, dd), s
        except ValueError:
            return None, s
    # "10–11/jul", "1–5/jul", "6–9/jul" -> pega o ÚLTIMO dia + mês abreviado
    m = re.search(r"(\d{1,2})\s*[–\-]\s*(\d{1,2})\s*/\s*([a-z]{3})", s.lower())
    if m and m.group(3) in MESES:
        dd = int(m.group(2)); mm = MESES[m.group(3)]
        try:
            return dt.datetime(2026, mm, dd), s
        except ValueError:
            return None, s
    # "dd/mês" textual: "27/mai"
    m = re.search(r"\b(\d{1,2})\s*/\s*([a-z]{3,9})\b", s.lower())
    if m and m.group(2) in MESES:
        dd = int(m.group(1)); mm = MESES[m.group(2)]
        try:
            return dt.datetime(2026, mm, dd), s
        except ValueError:
            return None, s
    # só nome de mês -> sem dia: deixa obs, sem data exata
    return None, s


def clean_status(val):
    if val is None:
        return "Não Iniciado"
    s = str(val).strip()
    low = s.lower()
    if low in ("em andamento",): return "Em Andamento"
    if low in ("não iniciado","nao iniciado"): return "Não Iniciado"
    if low in ("concluído","concluido","feito","finalizado","ok"): return "Concluído"
    if low in ("entregável","entregavel"): return "Entregável"
    if low in ("reunião","reuniao"): return "Reunião"
    return s


def parse_sheet(path, sheet, projeto, fonte, header_row=2, presencial_idx=6,
                detalhe_idx=2, resp_idx=3, prazo_idx=4, status_idx=5,
                skip_legend=True):
    """Lê uma aba template e devolve lista de dicts no schema do master.
    Colunas padrão: 0=Etapa,1=Atividade,2=Detalhe,3=Responsável,4=Prazo,5=Status,6=Presencial?
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    out = []
    sub_atual = None
    for i, row in enumerate(rows):
        if i <= header_row:
            continue
        row = list(row) + [None] * (8 - len(row)) if len(row) < 8 else list(row)
        etapa = row[0]
        atividade = row[1]
        # linha de legenda ("Legenda de Status:") ou cabeçalho repetido ("Ação")
        if isinstance(etapa, str) and etapa.strip().lower().startswith("legenda"):
            continue
        if isinstance(atividade, str) and atividade.strip().lower() in ("ação","acao","atividade"):
            # pode trazer um sub-projeto novo em col0
            if isinstance(etapa, str) and etapa.strip():
                sub_atual = etapa.strip()
            continue
        # atualiza sub-projeto (forward-fill)
        if isinstance(etapa, str) and etapa.strip():
            sub_atual = etapa.strip()
        # linha sem atividade real -> ignora (mas pode ser só um header de etapa)
        if not (isinstance(atividade, str) and atividade.strip()):
            continue
        detalhe = row[detalhe_idx] if detalhe_idx is not None and detalhe_idx < len(row) else None
        detalhe = str(detalhe).strip() if isinstance(detalhe, str) and detalhe.strip() else None
        resp_raw = row[resp_idx] if resp_idx < len(row) else None
        resp_list, resp_raw_s = split_responsaveis(resp_raw)
        prazo, prazo_obs = parse_prazo(row[prazo_idx] if prazo_idx < len(row) else None)
        status = clean_status(row[status_idx] if status_idx < len(row) else None)
        eh_entregavel = (status == "Entregável") or bool(
            re.search(r"entreg|entregável|versão final|lançamento|publica", str(atividade).lower())
        )
        presencial = None
        if presencial_idx is not None and presencial_idx < len(row):
            pv = row[presencial_idx]
            if isinstance(pv, str) and pv.strip():
                lp = pv.strip().lower()
                if lp.startswith("remoto"): presencial = "Remoto"
                elif lp.startswith("presenc"): presencial = "Presencial"
                elif lp in ("sim","ok","x"): presencial = "Presencial"
        out.append({
            "projeto": projeto,
            "sub_projeto": sub_atual,
            "atividade": str(atividade).strip(),
            "detalhe": detalhe,
            "responsaveis": "; ".join(resp_list),
            "responsaveis_raw": resp_raw_s,
            "n_responsaveis": len(resp_list),
            "prazo": prazo,
            "prazo_obs": prazo_obs,
            "status": status,
            "eh_entregavel": eh_entregavel,
            "esforco": None,
            "fonte": fonte,
            "presencial": presencial,
        })
    return out


if __name__ == "__main__":
    rows = parse_sheet("Cronograma_Comunicação_2026.xlsx", "Modelo", "Comunicação",
                       "Comunicação", header_row=2)
    for r in rows:
        print(r["sub_projeto"], "|", r["atividade"][:40], "|", r["responsaveis"],
              "|", r["prazo"], "|", r["status"], "|ent=", r["eh_entregavel"])
    print("TOTAL:", len(rows))
