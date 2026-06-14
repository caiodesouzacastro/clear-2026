# Atualização do dashboard — Aba "Alocação" + correções

## O que mudou nesta versão
1. **Nova aba "Alocação"** (entre "Calendário" e "Equipe") com dois Gantts:
   - **Por pessoa** — cada pessoa e as frentes em que está, mês a mês (cor = projeto).
   - **Por projeto** — cada frente e quem está nela (cor = pessoa).
   - Filtro de meses (slider), tooltips ao passar o mouse e cartões de **pico de
     frentes simultâneas** por pessoa.
   - Linhas divisórias reforçadas + faixa de fundo alternada para separar bem
     cada pessoa / cada projeto.
2. **Nomes corrigidos em TODO o master:** Pleno → **Fred**, Senior → **Hisrael**,
   Zinho → **Michel** (inclusive em observações e textos livres da aba Atividades).
3. **Nova aba no master: `Alocacao_Gantt`** (pessoa, projeto, mes_de, mes_ate, origem),
   que alimenta a aba Alocação. Fonte = `team_allocation.csv` (frentes novas de IU, TCE,
   Rede Subnacional, AfrEA — vencem em conflito) **+** o Envolvimento do master
   (PRIME, EVALAC, REDECA, Bens Públicos, FINEP, etc.).

## Como colocar no ar (passo a passo)
1. No repositório `caiodesouzacastro/clear-2026`, substitua **2 arquivos** pelos novos:
   - `app.py`
   - `CLEAR_Master_2026.xlsx`
2. Faça o commit e o push:
   ```bash
   git add app.py CLEAR_Master_2026.xlsx
   git commit -m "Nova aba Alocacao + Fred/Hisrael + Zinho->Michel"
   git push
   ```
3. O Streamlit Cloud redeploya sozinho em ~1-2 min. Se o site não atualizar,
   abra share.streamlit.io e clique em **Reboot** (ou faça um commit qualquer no app.py).

`requirements.txt` **não muda** (usa só `streamlit.components`, que já vem no streamlit).

## Atualizações futuras da alocação
Gere um novo `team_allocation.csv` (colunas: Project, Project start, Project end,
Person, Person from, Person to — meses como números, 6 = Junho) e rode:
```bash
python3 regen_master.py
```
Isso refaz o `CLEAR_Master_2026.xlsx` já com os nomes corrigidos e a aba Alocacao_Gantt.
