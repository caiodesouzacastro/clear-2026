# Atualizar o Dashboard CLEAR 2026 — Guia + Prompt Único

**Repositório:** `caiodesouzacastro/clear-2026` (público)
**Dashboard:** `https://clear-2026-ahhtunccmwdbxx5nmfxpbb.streamlit.app`
**Última atualização deste doc:** junho 2026

> Este documento é o **único** ponto de referência para atualizar o dashboard. Substitui o
> `COMO_ATUALIZAR.md` e o `PROMPT_gerar_planilha.md` antigos. Tem duas partes:
> **Parte A** — guia rápido (não-técnico) de como subir e atualizar.
> **Parte B** — o prompt completo de regeneração (cole no Claude com os cronogramas anexados).

---

# PARTE A — Guia rápido

## O que existe hoje

**Dashboard — 7 abas:**

| Aba | O que mostra |
|---|---|
| Visão Geral | Gantt por projeto com % concluído, filtro de projetos, marcadores de entregáveis |
| Calendário | Grade mensal; itens de **Comunicação** saem em **âmbar** com legenda própria |
| Comunicação | Painel dedicado: todas as atividades de Comunicação, sem depender do filtro lateral |
| Alocação | Gantt por pessoa ou por projeto, com filtros |
| Equipe | Heatmap de horas por pessoa × mês |
| Pesquisador | Visão individual: carga + lista de atividades |
| Farol | Percepção de carga por pessoa × mês (Esperada × Percebida) |

**Arquivos no repositório:** `app.py` (código), `CLEAR_Master_2026.xlsx` (dados),
`Farol_Intensidade_CLEAR_2026.xlsx` (percepção da equipe), `requirements.txt` e `config.toml`
(não mudam).

**Arquivos de apoio (regeneração):** `parser.py`, `regen_master_v2.py` e este prompt.

## Como subir uma atualização

1. Acesse `github.com/caiodesouzacastro/clear-2026`.
2. Substitua os arquivos alterados (`CLEAR_Master_2026.xlsx`, `Farol_Intensidade_CLEAR_2026.xlsx`
   e/ou `app.py`).
3. **Commit changes** (botão verde).
4. Aguarde 1–2 min — o Streamlit redeploya sozinho.
5. Se não atualizar: `share.streamlit.io` → app `clear-2026` → **Reboot**.

> `requirements.txt` não muda. Mudanças no `app.py` (cor/aba da Comunicação, ordem das abas)
> já estão no código — **não** se regeneram a cada ciclo. O ciclo normal mexe só nos **dados**.

## Ciclo de atualização dos dados

1. Projeto preenche o `Modelo_Cronograma_CLEAR_2026.xlsx` (aba **Modelo**).
2. Junte todos os cronogramas + este prompt + `parser.py`/`regen_master_v2.py` e mande pro Claude.
3. Claude gera o `CLEAR_Master_2026.xlsx` novo seguindo a Parte B.
4. Suba no GitHub (passos acima). Se veio Farol novo, suba junto.

---

# PARTE B — Prompt de regeneração (cole no Claude)

## 1. Objetivo

Regenerar o `CLEAR_Master_2026.xlsx` a partir dos cronogramas anexados, **preservando** as abas
que não vêm dos cronogramas e aplicando as regras de normalização e a taxonomia abaixo.

## 2. O que você recebe

- **N cronogramas** (.xlsx) no template padrão.
- **O master atual** — para copiar abas preservadas e o EVALAC.
- (opcional) **Deck da plenária ReDeCA** (.pptx) — marcos da ReDeCA.
- (opcional) **`Farol_Intensidade_CLEAR_2026.xlsx`** novo — percepção da equipe.

## 3. Formato do template de cronograma

- Linha 0 título · Linha 1 descrição · Linha 2 cabeçalho.
- Cabeçalho: `Etapa | Atividade | [Detalhe] | Responsável | Prazo | Status | [Presencial?/Formato]`.
  A coluna **Detalhe** e a **Presencial?** podem não existir — confira os índices em cada arquivo.
- Linha "Legenda de Status:" logo abaixo do cabeçalho → ignorar.
- Coluna **Etapa** vem mesclada: *forward-fill* e use como `sub_projeto`.
- **Inspecione cada arquivo**: os índices de coluna variam.

## 4. Normalização de nomes (SEMPRE, em todas as células de todas as abas)

| De | Para | Motivo |
|---|---|---|
| Luiggi, Luigui | **Luigi** | um "g" só |
| Zinho | **Michel** | artefato; "Zinho" não existe |
| Pleno | **Fred** | apelido de papel |
| Senior, Sênior | **Hisrael** | apelido de papel |
| Samy | **Samu** | variante |
| Beatriz B, Bai B | **Bia B** | variante |
| **BiaS, BiaB** (sem espaço) | **Bia S / Bia B** | variante grudada |
| Ju, Júlia | **Julia** | variante |
| Fabricia | **Fabrícia** | acento |
| Cecília | **Cecilia** | canônico sem acento |
| Júnior | **Junior** | canônico sem acento |
| Caraol | **Carol** | typo |
| PMO | *(descartar)* | não é pessoa |
| Bia *(sozinho)* | *(descartar)* | ambíguo |

**Time canônico (16):** Bia B, Bia S, Caio, Carol, Cecilia, Fabrícia, Fred, Hisrael, Julia,
Junior, Lorena, Luan, Luigi, Lycia, Michel, Samu.

**Externos reais a preservar:** Marcio, Lara, Lia, Amanda, Alei, Hugo, Gustavo (gus→Gustavo),
Paulinha, Bárbara, Livia, Tania, Cristina Navas, Gabriel W, Isabella, Ana, Auditores, IFAD, OVE,
BID, OVE/BID, Sicredi, TCE, Comitê Gestor, Aulasneo, Cross Content, CLEAR, CLEAR LAC, CLEAR LAB,
Rede CLEAR, ONU Mulheres, AfrEA, TI FGV, Biblioteca, Equipe Implementadora, CMF-BID, Comunicação FGV.

**Descartar** notas de supervisão ("com ajuda/apoio", "supervisionada por", "sup.", "Cc", "Obs:"),
parênteses soltos e qualquer nome fora das listas. Não inventar pessoas.

## 5. Parsing das atividades

- **Datas:** serial Excel (converter!), `dd/mm/aaaa`, "22 de junho", "10–11/jul" (último dia),
  "15/jul". Mês sem dia → `prazo` vazio, texto em `prazo_obs`. Texto não-data → `prazo_obs`.
- **`responsaveis`:** separar por vírgula / "e" / "/" / "+"; limpar; manter só canônicos+externos;
  guardar bruto em `responsaveis_raw` e contagem em `n_responsaveis`.
- **`eh_entregavel`:** `True` se `status=="Entregável"` ou atividade casa `entreg|versão final|lançamento|publica`.
- **`status`:** mapear para `{Concluído, Em Andamento, Não Iniciado, Atrasado, Entregável, Reunião}`.

## 6. Taxonomia de projetos e decisões firmadas

- **Comunicação** = projeto próprio. Nunca fundir.
- **ILUMA - Rede Lusófona** = projeto. **Substituiu o ILUSOMA, que foi aposentado** — não preservar ILUSOMA.
- **Bens Públicos** = projeto consolidado. Inclui o **MiniGuia** como **sub-projeto** (`sub_projeto="MiniGuia"`),
  além de Guias / Sínteses / Avaliações Executivas. **MiniGuia não é projeto próprio.**
- **Rede PARES** = **projeto próprio** (novo; rede de pares, fases de Preparação/Lançamento/Mobilização).
- **Painel CLEAR** = projeto (arquivo de Monitoramento).
- **PRiME III** = juntar abas "PRiME" + "PRiME III IE — Revisões".
- **EVALAC** = preservar do master antigo se não vier cronograma novo dele.
- **ReDeCA** = marcos do deck da plenária; **Caio é o único responsável**.

**Mapa cronograma → projeto → aba → colunas** (0-based; cabeçalho na linha 2; "—" = coluna ausente):

| Projeto | Arquivo | Aba | detalhe | resp | prazo | status | presencial |
|---|---|---|---|---|---|---|---|
| TCE ES | TCE_Cronograma | "TCE ES " | 2 | 3 | 4 | 5 | — |
| PRiME III | PRiME_Cronograma | "PRiME" | 2 | 3 | 4 | 5 | 6 |
| PRiME III | PRiME_Cronograma | "PRiME III IE — Revisões" | — | 2 | 3 | 4 | — |
| IU - Oficina Sistemas | IU_OficinaSistemas | "Oficina Sistemas_IU" | 2 | 3 | 4 | 5 | — |
| Sicredi | IU_OficinaSistemas | "Sicredi" | 2 | 3 | 4 | 5 | — |
| Bens Públicos | BensPúblicos | "Bens Públicos 2026" | — | 2 | 3 | 4 | 5 |
| Bens Públicos (sub MiniGuia) | Miniguia_Cronograma | "Miniguia" | — | 2 | 3 | 4 | 5 |
| Comunicação | Comunicação | "Modelo" | 2 | 3 | 4 | 5 | 6 |
| ILUMA - Rede Lusófona | ILUMA_RedeLusofona | "Modelo" | 2 | 3 | 4 | 5 | — |
| Rede PARES | Rede_PARES | "Modelo" | 2 | 3 | 4 | 5 | 6 |
| Painel CLEAR | PainelMonitoramento | "Painel CLEAR" | — | 2 | 3 | 4 | 5 |

> **MiniGuia:** parsear com `projeto="Bens Públicos"`, depois sobrescrever `sub_projeto="MiniGuia"`
> (a etapa original — Planejamento / Elaboração — vai para o `detalhe`).

## 7. ReDeCA — marcos do deck (projeto "ReDeCA", responsável **Caio**, fonte "Plenária ReDeCA 14/05/2026")

1. **14/05/2026** — Sessão plenária + envio pós-plenária/guia/instrumento — Concluído · entregável
2. **27/05/2026** — Encerramento das respostas à pós-plenária — Concluído · entregável
3. **22/06/2026** — Encerramento da pré-fase / fim da coleta documental — Em Andamento · entregável
4. **Jul–Set 2026** — Implementação em 3 blocos (sem data fixa → `prazo_obs`) — Não Iniciado
5. **15/07/2026** (obs "Julho 2026") — Café ReDeCA com a ONU Mulheres — Não Iniciado · entregável
6. **15/10/2026** (obs "Outubro 2026 — Lima") — Encontro Anual em Lima, 29 perfis — Não Iniciado · entregável

## 8. Abas a PRESERVAR sem tocar

Copie do master antigo, byte a byte: `Alocacao`, `Envolvimento`, `Alocacao_Gantt`, `Carga_Esperada`.

> O Farol lê de `Carga_Esperada` (taxonomia própria, CAIXA ALTA, mais granular — não bate 1:1
> com os nomes de projeto das Atividades, e isso é proposital) + do arquivo à parte
> `Farol_Intensidade_CLEAR_2026.xlsx`. **Não alterar** pela regeneração.

## 9. Farol de percepção (drop-in)

Se vier um `Farol_Intensidade_CLEAR_2026.xlsx` novo, só substituir o arquivo no repo.
Aba "Farol de Intensidade", cabeçalho na linha 4, uma linha por pessoa, colunas Jan–Dez,
escala **0–3** (0 sem dados · 1 tranquilo · 2 ok · 3 sobrecarga), e linha final "TOTAL EQUIPE".

## 10. Schema do master (9 abas)

- **README** — atualizar "Última atualização" para hoje.
- **Atividades** — `id_atividade, projeto, sub_projeto, atividade, detalhe, responsaveis,
  responsaveis_raw, n_responsaveis, prazo, prazo_obs, status, eh_entregavel, esforco, fonte, presencial`.
- **Resp_Atividade** — uma linha por (atividade × responsável): `id_atividade, pessoa, projeto,
  prazo, status, eh_entregavel`.
- **Pessoas** — os 16 canônicos · **Projetos** — lista derivada das Atividades.
- **Alocacao, Envolvimento, Alocacao_Gantt, Carga_Esperada** — preservadas (seção 8).

## 11. Checklist de validação

- [ ] Zero vazamento de "Zinho", "Pleno", "Senior", "Luiggi", "BiaS" em qualquer aba.
- [ ] `Resp_Atividade` só com canônicos + externos reais.
- [ ] `Pessoas` = os 16 canônicos · `id_atividade` único · zero órfãos em Resp.
- [ ] `prazo` é datetime; seriais Excel convertidos.
- [ ] Abas de alocação idênticas ao master antigo · README com data de hoje.
- [ ] ILUSOMA **não** aparece · EVALAC preservado · MiniGuia dentro de Bens Públicos ·
      Rede PARES como projeto próprio.

## 12. Saída e deploy

Gerar o `CLEAR_Master_2026.xlsx`, entregar para download, subir no GitHub (`main`),
aguardar ~2 min, Reboot se necessário. Se houve Farol novo, subir junto.
