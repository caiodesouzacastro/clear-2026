# Guia de Atualização — Dashboard CLEAR 2026

**Repositório:** `caiodesouzacastro/clear-2026`
**Dashboard:** `https://clear-2026-ahhtunccmwdbxx5nmfxpbb.streamlit.app`
**Última atualização deste guia:** junho 2026

---

## O que existe hoje

### Dashboard (site)
6 abas:

| Aba | O que mostra |
|---|---|
| **Visão Geral** | Gantt por projeto com % concluído, filtro de projetos, marcadores de entregáveis |
| **Calendário** | Grade mensal com todas as atividades e entregáveis do mês |
| **Alocação** | Gantt por pessoa ou por projeto (quem está em quê, e quando) — com filtros de pessoa e projeto |
| **Farol** | Heatmap pessoa × mês mostrando intensidade de carga (verde / âmbar / vermelho) |
| **Equipe** | Heatmap de horas por pessoa × mês |
| **Pesquisador** | Visão individual: gráfico de carga + lista de atividades |

### Arquivos no repositório
- `app.py` — código do dashboard
- `CLEAR_Master_2026.xlsx` — planilha mestra com 8 abas internas
- `requirements.txt` — dependências (não muda)
- `config.toml` — tema visual (não muda)

### Arquivos de uso interno (não vão pro repositório)
- `Modelo_Cronograma_CLEAR_2026.xlsx` — template para novos projetos
- `Farol_Intensidade_CLEAR_2026.xlsx` — auto-declaração de carga por pessoa
- `regen_master_v2.py` — script que regenera o master a partir dos cronogramas

---

## Como colocar uma atualização no ar

São sempre **2 arquivos** que vão pro GitHub:

1. Acesse `github.com/caiodesouzacastro/clear-2026`
2. Substitua `app.py` e/ou `CLEAR_Master_2026.xlsx` pelos novos
3. Desça até **"Commit changes"** e clique no botão verde
4. Aguarde 1–2 minutos — o Streamlit redeploya automaticamente
5. Se o site não atualizar: entre em `share.streamlit.io`, ache o app `clear-2026` e clique em **Reboot**

> `requirements.txt` **não muda.**

---

## Como atualizar o master quando chegar novo cronograma

1. Peça ao projeto que preencha o `Modelo_Cronograma_CLEAR_2026.xlsx`
   (aba **Modelo** em branco; a aba EVALAC 2026 é exemplo de preenchimento)
2. Junte todos os arquivos de cronograma e suba junto com o `regen_master_v2.py`
   para o Claude (ou rode localmente):
   ```bash
   python3 regen_master_v2.py
   ```
3. O script gera um novo `CLEAR_Master_2026.xlsx` já com:
   - Todos os cronogramas consolidados
   - Nomes normalizados (Zinho→Michel, Luigi→Luiggi, etc.)
   - Aba `Alocacao_Gantt` atualizada
4. Sobe o novo master pro GitHub conforme o passo anterior

---

## Regras de normalização de nomes (sempre aplicadas)

| De | Para |
|---|---|
| Zinho | Michel |
| Pleno | Fred |
| Senior | Hisrael |
| Luigi | Luiggi |

> Essas trocas acontecem em **todas** as células de **todas** as abas, inclusive textos livres e observações.

---

## Estrutura interna do CLEAR_Master_2026.xlsx

| Aba | Conteúdo |
|---|---|
| `README` | Metadados e data da última geração |
| `Atividades` | Cronograma consolidado (424 atividades, 9 projetos) |
| `Resp_Atividade` | Relação pessoa × atividade (expandida por responsável) |
| `Alocacao` | Horas/mês legado (mantido para compatibilidade) |
| `Envolvimento` | Envolvimento mensal por pessoa × projeto |
| `Alocacao_Gantt` | De–até por pessoa × projeto (alimenta aba Alocação e Farol) |
| `Pessoas` | Lista canônica da equipe |
| `Projetos` | Lista canônica de projetos |
