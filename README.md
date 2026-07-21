# CLEAR 2026 — Dashboards de Organização, Alocação e Dependências

Este repositório hospeda **dois apps Streamlit** que leem os mesmos arquivos de dados:

## 🔗 Acesse os painéis

**📊 Organização & Alocação**
**👉 [https://clear2026-cronograma.streamlit.app](https://clear2026-cronograma.streamlit.app)**
Portfólio de projetos, alocação de pesquisadores e próximas entregas do FGV CLEAR em 2026.

**🔗 Cadeia de Dependências**
**👉 [https://clear2026-dependencias.streamlit.app](https://clear2026-dependencias.streamlit.app)**
Fio condutor entre atividades: o que trava o quê, dentro de um projeto e entre projetos.

---

## Estrutura do repositório

```
clear-2026/
├── app.py                          # App 1 — Organização & Alocação
├── app_dependencias.py             # App 2 — Cadeia de Dependências
├── CLEAR_Master_2026.xlsx          # Fonte única de dados (editável)
├── Dependencias_CLEAR_2026.xlsx    # Arestas declaradas do grafo (editável)
├── config.toml                     # Tema visual
├── requirements.txt
└── README.md
```

A planilha mestra é a **única fonte de dados** dos dois apps. Editou → os apps refletem em 1-2 minutos.

---

## App 1 — Organização & Alocação (`app.py`)

### 📊 Visão Geral
Linha do tempo (Gantt) por projeto, marcadores de entregáveis e KPIs. *Como está distribuído o ano?*

### 📅 Calendário
Grade mensal navegável com as entregas do mês. *O que acontece este mês?*

### 👥 Equipe
Mapa de carga — pessoas × meses, intensidade pelo número de projetos simultâneos. *Quem está sobrecarregado?*

### 👤 Pesquisador
A fundo numa pessoa: capacidade, projetos simultâneos, atividades atribuídas. *Qual é a carga desta pessoa?*

**Filtros (sidebar):** só entregáveis críticos ou todas as atividades · subset de projetos.

---

## App 2 — Cadeia de Dependências (`app_dependencias.py`)

Lê o `CLEAR_Master_2026.xlsx` (status/prazo ao vivo) e o `Dependencias_CLEAR_2026.xlsx` (o grafo).
Convenção: aresta **A → B** lê-se *"A destrava B"* (B depende de A).

### ⬇️ Impacto a jusante
Escolhe uma atividade e ilumina toda a cadeia que depende dela, até as linhas de chegada atingidas. *Se isto atrasar, o que cai junto?*

### 🎯 Caminho crítico por linha de chegada
A corrente mais longa que sustenta cada marco. *O que determina esta data?*

### 🔀 Mapa inter-projetos
Só as arestas que cruzam projeto — as mais caras quando quebram.

**Legenda:** cor do nó = status atual no Master · linha cheia = dependência dura · tracejada = mole · pontilhada roxa = precedência provável (derivada, só Painel/EVALAC).

### Como o grafo é alimentado (`Dependencias_CLEAR_2026.xlsx`)
- **Aba `Nos`** — cada entregável tem um `dep_id` **estável** (ex.: `PAINEL-D03`). O app re-encontra a atividade no Master pela chave `(projeto, atividade, detalhe)`; se um rename quebrar a âncora, o painel acende um **alerta de nó órfão**.
- **Aba `Arestas`** — uma linha por dependência: `origem_id → destino_id`, `tipo` (bloqueio/insumo), `forca` (dura/mole).
- **Aba `Flags`** — pendências conhecidas (gaps de nó, datas a corrigir).

Para adicionar dependências, basta editar essa planilha e subir — o app não muda.

---

## Como atualizar os dados

1. Cada projeto preenche seu cronograma (`Modelo_Cronograma_CLEAR_2026.xlsx`).
2. Uma pessoa junta todos os cronogramas no Claude usando `PROMPT_ATUALIZAR_CLEAR_2026.md` e gera o novo `CLEAR_Master_2026.xlsx`.
3. Sobe o arquivo no GitHub substituindo o atual — os painéis atualizam sozinhos.

---

## 16 pessoas CLEAR (nomes canônicos)

Bia B, Bia S, Caio, Carol, Cecilia, Fabrícia, Fred, Hisrael, Julia, Junior, Lorena, Luan, Luigi, Lycia, Michel, Samu

---

## Rodar localmente (dev)

```bash
git clone https://github.com/caiodesouzacastro/clear-2026
cd clear-2026
pip install -r requirements.txt

streamlit run app.py                 # App 1
streamlit run app_dependencias.py    # App 2
```

Abre em `http://localhost:8501`.
