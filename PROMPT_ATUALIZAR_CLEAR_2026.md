# Atualizar o Dashboard CLEAR 2026 — Guia + Prompt Único

**Repositório:** `caiodesouzacastro/clear-2026` (público)
**Dashboard:** `https://clear-2026-ahhtunccmwdbxx5nmfxpbb.streamlit.app`
**Última atualização deste doc:** 12/10/2026

> Este documento é o **único** ponto de referência para atualizar o dashboard.
> **Parte A** é o guia rápido de como subir. **Parte B** é o prompt de regeneração
> (cole no Claude com os cronogramas anexados).

---

# PARTE A — Guia rápido

## O que existe hoje

**Dashboard, 10 abas:**

| Aba | O que mostra |
|---|---|
| Visão Geral | Gantt por projeto com % concluído, filtro de projetos, marcadores de entregáveis |
| Calendário | Grade mensal; itens de **Comunicação** saem em **âmbar** com legenda própria |
| Comunicação | Painel dedicado: todas as atividades de Comunicação, sem depender do filtro lateral |
| Eventos | Eventos do centro (webinars, lançamentos, mesas, encontros), com próximo evento em destaque |
| Alocação | Gantt por pessoa ou por projeto, com filtros |
| Equipe | Heatmap de horas por pessoa × mês |
| Pesquisador | Visão individual: carga + lista de atividades |
| Farol | Percepção de carga por pessoa × mês (Esperada × Percebida) |
| Férias | Linha do tempo de férias, com sobreposições sombreadas |
| Monitorias | Sessões de mentoria da Formação CPP 1.2, por mentor, rodada e grupo |

**Dados no repositório:** `CLEAR_Master_2026.xlsx` (master), `Farol_Intensidade_CLEAR_2026.xlsx`,
`Ferias_CLEAR_2026.xlsx`, `Dependencias_CLEAR_2026.xlsx`.
**Código:** `app.py`, `requirements.txt`, `config.toml` (os dois últimos não mudam).
**Pipeline de regeneração:** `regen_master_v4.py`, `parser.py`, `gantt.py`, `finalizar.py`.

> `regen_master.py` e `regen_master_v2.py` são versões antigas e podem ser apagadas.

## Como subir uma atualização

1. Acesse `github.com/caiodesouzacastro/clear-2026`.
2. Substitua os arquivos alterados.
3. **Commit changes** (botão verde).
4. Aguarde 1 ou 2 minutos; o Streamlit redeploya sozinho.
5. Se não atualizar: `share.streamlit.io` → app `clear-2026` → **Reboot**.

O ciclo normal mexe só nos **dados**. Mudanças no `app.py` não se regeneram a cada ciclo.

## Ciclo de atualização dos dados

1. Cada projeto preenche o `Modelo_Cronograma_CLEAR_2026.xlsx` (aba **Modelo**).
2. Junte todos os cronogramas e mande pro Claude com este prompt.
3. Claude baixa os arquivos vivos do GitHub, regenera o master e devolve pra download.
4. Suba no GitHub (passos acima).

## Como rodar o pipeline na mão

```bash
export CLEAR_CRONOGRAMAS=./cronogramas        # pasta com os Cronograma_*.xlsx
export CLEAR_MASTER_ANTIGO=./CLEAR_Master_ANTIGO.xlsx   # master vivo do GitHub
export CLEAR_OUT=./CLEAR_Master_2026.xlsx
python regen_master_v4.py && python finalizar.py
```

`regen_master_v4.py` parseia os cronogramas e monta Atividades, Resp_Atividade,
Pessoas e Projetos. `finalizar.py` faz o resto: correções pontuais, Sicredi,
caixa dos status, `Alocacao_Gantt`, `Carga_Esperada`, `Eventos` e `Monitorias`.

---

# PARTE B — Prompt de regeneração (cole no Claude)

## 1. Objetivo

Regenerar o `CLEAR_Master_2026.xlsx` a partir dos cronogramas anexados, **preservando**
as abas que não vêm dos cronogramas e aplicando as regras abaixo.

## 2. Autoridade de versão (ler antes de tudo)

- **Sempre baixe os arquivos vivos do GitHub antes de regenerar.** O zip do OneDrive
  costuma estar atrasado. Use `https://raw.githubusercontent.com/caiodesouzacastro/clear-2026/main/{arquivo}?nocache=$(date +%s)`.
- Para comparar versões use o `dcterms:modified` de `docProps/core.xml` dentro do xlsx,
  não o timestamp do zip nem do sistema de arquivos.
- **Nunca sobrescreva `Farol_Intensidade_CLEAR_2026.xlsx` nem `Ferias_CLEAR_2026.xlsx`
  sem conferir qual é mais novo.** O Farol do zip já veio atrasado duas vezes seguidas,
  e sobrescrever apaga percepção que a equipe preencheu à mão.

## 3. Formato do template de cronograma

- Linha 0 título, linha 1 descrição, linha 2 cabeçalho.
- Cabeçalho: `Etapa | Atividade | [Detalhe] | Responsável | Prazo | Status | [Presencial?]`.
- Linha "Legenda de Status:" logo abaixo do cabeçalho: ignorar.
- Coluna **Etapa** vem mesclada: *forward-fill*, e virá o `sub_projeto`.
- **Inspecione cada arquivo**: os índices de coluna e a linha do cabeçalho variam.
  O CPP 1.2 tem cabeçalho na linha 1, não na 2.

## 4. Normalização de nomes (em todas as células de todas as abas)

| De | Para |
|---|---|
| Luiggi, Luigui | **Luigi** (um "g" só) |
| Zinho | **Michel** |
| Pleno | **Fred** |
| Senior, Sênior, Frederik, Frederik de Jonghe | **Fred** é Frederik; **Hisrael** é o Senior |
| Samy, Samuely | **Samu** |
| Beatriz B, Bai B, BiaB | **Bia B** |
| BiaS | **Bia S** |
| Ju, Júlia | **Julia** |
| Fabricia | **Fabrícia** |
| Cecília | **Cecilia** · Júnior → **Junior** (canônicos sem acento) |
| Caraol | **Carol** |
| PMO, Bia (sozinho) | *descartar* |

> Atenção: `Senior` e `Sênior` mapeiam para **Hisrael**, `Pleno` e `Frederik` para **Fred**.

**Time canônico (16):** Bia B, Bia S, Caio, Carol, Cecilia, Fabrícia, Fred, Hisrael,
Julia, Junior, Lorena, Luan, Luigi, Lycia, Michel, Samu.

**Externos reais a preservar:** Marcio, Lara, Lia, Amanda (Amaninha), Alei, Hugo,
Gustavo (gus), Paulinha, Bárbara, Livia, Tania, Cristina Navas, Gabriel W (Gabriel),
Cenci, Khalid, Thania, Isabella, Ana, Thais, Fabiana, Fábio, Clóvis, Auditores, IFAD,
OVE, BID, OVE/BID, Sicredi, TCE, Comitê Gestor, Aulasneo, Cross Content, CLEAR,
CLEAR LAC, CLEAR LAB (FGV CLEAR LAB), Rede CLEAR, ONU Mulheres, AfrEA, TI FGV,
Biblioteca, Equipe Implementadora, CMF-BID, Comunicação FGV.

**Descartar:** notas de supervisão ("com ajuda/apoio", "supervisionada por", "sup.",
"Cc", "Obs:"), parênteses soltos, e placeholders de papel que não são pessoas
("A definir", "TBD", "Mentores", "Equipe", "Consultant", "Junior person",
"Participantes", "Hands-on trainer", "gestores", "—"). Não inventar pessoas.

Dois padrões que o `parser.py` já trata:
- `"Fábio: dados secundários"`: fica só o nome antes dos dois-pontos.
- `"Samu Michel e Julia"` (vírgula faltando na origem): o `finalizar.py` expande.
  Vale pedir pro responsável do cronograma colocar a vírgula.

## 5. Parsing das atividades

- **Datas:** serial Excel (converter), `dd/mm/aaaa`, "22 de junho", "10–11/jul"
  (último dia), "15/jul". Mês sem dia: `prazo` vazio e o texto vai pro `prazo_obs`.
- **Confira anos fora da faixa 2025–2028**: já apareceu `207-01-22` em vez de 2027.
  Correções pontuais ficam no dict `CORRECOES_PRAZO` do `finalizar.py`.
- **`responsaveis`:** separar por vírgula, "e", "/", "+"; guardar bruto em
  `responsaveis_raw` e a contagem em `n_responsaveis`.
- **`eh_entregavel`:** `True` se `status=="Entregável"` ou a atividade casa
  `entreg|versão final|lançamento|publica`.
- **`status`:** só estes 5 valores, porque o `app.py` compara string exata:
  `{Concluído, Em Andamento, Não Iniciado, Atrasado, Reunião}`. Qualquer outra coisa
  precisa de mapeamento no `STATUS_MAP` do `finalizar.py` (já cobre "atrasado"
  minúsculo, "pronto" e "Realizado parcialmente").

## 6. Taxonomia de projetos e decisões firmadas

- **Comunicação** é projeto próprio. Nunca fundir.
- **ILUMA - Rede Lusófona** substituiu o ILUSOMA, que foi aposentado. Não preservar ILUSOMA.
- **Bens Públicos** inclui o **MiniGuia** como sub-projeto (`sub_projeto="MiniGuia"`).
  MiniGuia não é projeto próprio.
- **Rede PARES** é projeto próprio.
- **Painel CLEAR** é projeto (arquivo de Monitoramento).
- **PRiME III** vem só da aba `PRiME`. A aba `PRiME III IE — Revisões` foi apagada
  do cronograma em outubro/2026 e **não é mais preservada**.
- A aba **`Hands-on`** do cronograma do PRiME **não é ingerida**: em 12/10/2026, 44 das
  54 atividades dela duplicavam a aba principal, com status em inglês
  ("Concluded", "ongoing", "not started"), e as 10 restantes eram a versão antiga do
  ciclo de revisão do workbook, já substituída na aba principal. É cópia de trabalho.
- As abas `Entregas - fornecedores`, `Produtos - PRiME`, `Produtos - TCE ES`,
  `Planilha1` (decisões com o IFAD), `Parâmetros`, `Mentores` e `Visão por data`
  também não são ingeridas.
- **EVALAC** é preservado do master antigo se não vier cronograma novo dele.
- **ReDeCA** vem dos marcos da plenária; **Caio é o único responsável**.

**Mapa cronograma → projeto → aba → colunas** (0-based; "—" = coluna ausente):

| Projeto | Arquivo | Aba | cabeçalho | detalhe | resp | prazo | status | presencial |
|---|---|---|---|---|---|---|---|---|
| TCE ES | Cronograma_TCE | "TCE ES " | 2 | 2 | 3 | 4 | 5 | — |
| PRiME III | Cronograma_PRiME | "PRiME" | 2 | 2 | 3 | 4 | 5 | 6 |
| IU - Oficina Sistemas | IU_OficinaSistemas | "Oficina Sistemas_IU" | 2 | 2 | 3 | 4 | 5 | — |
| Sicredi | Cronograma_Sicredi_20260406 | "Sicredi" | 2 | 2 | 3 | 4 | 5 | — |
| Bens Públicos | BensPúblicos | "Bens Públicos 2026" | 2 | — | 2 | 3 | 4 | 5 |
| Bens Públicos (sub MiniGuia) | Cronograma_Miniguia | "Miniguia" | 2 | — | 2 | 3 | 4 | 5 |
| Comunicação | Comunicação | "Modelo" | 2 | 2 | 3 | 4 | 5 | 6 |
| ILUMA - Rede Lusófona | ILUMA_RedeLusofona | "Modelo" | 2 | 2 | 3 | 4 | 5 | — |
| Rede PARES | Rede_PARES | "Modelo" | 2 | 2 | 3 | 4 | 5 | 6 |
| Painel CLEAR | PainelMonitoramento | "Painel CLEAR" | 2 | — | 2 | 3 | 4 | 5 |
| IU - Assessoria Técnica | IU_AT_SEDU_ES | "Cronograma" | 2 | 2 | 3 | 4 | 5 | 6 |
| IU - Avaliação Executiva | Avaliação_Exec_2026 | "Planilha1" | 1 | 2 | 3 | 5 | 6 | 8 |
| IU - Formação CPP | IU_ CPP1.2 | "Cronograma" | **1** | 2 | 3 | 4 | 5 | — |

> **MiniGuia:** parsear com `projeto="Bens Públicos"` e depois sobrescrever
> `sub_projeto="MiniGuia"` (a etapa original vai para o `detalhe`).
> **IU - Avaliação Executiva:** fora do template, tem Prazo inicial e Prazo final,
> e linhas que só trazem a Etapa mas têm prazo viram atividade própria.

## 7. ReDeCA, marcos do deck (projeto "ReDeCA", responsável **Caio**, fonte "Plenária ReDeCA 14/05/2026")

1. **14/05/2026** Sessão plenária + envio pós-plenária/guia/instrumento. Concluído, entregável
2. **27/05/2026** Encerramento das respostas à pós-plenária. Concluído, entregável
3. **22/06/2026** Encerramento da pré-fase e fim da coleta documental. Em Andamento, entregável
4. **Jul a Set 2026** Implementação em 3 blocos (sem data fixa, vai pro `prazo_obs`). Não Iniciado
5. **15/07/2026** (obs "Julho 2026") Café ReDeCA com a ONU Mulheres. Não Iniciado, entregável
6. **15/10/2026** (obs "Outubro 2026, Lima") Encontro Anual em Lima, 29 perfis. Não Iniciado, entregável

## 8. Abas a PRESERVAR sem tocar

Copie do master antigo: `Alocacao` e `Envolvimento`.

`Alocacao_Gantt` e `Carga_Esperada` são **derivadas**, não preservadas:

- `Alocacao_Gantt` sai do `gantt.py`, a partir das Atividades. Taxonomia mais granular
  que `projeto` (PRiME em 7 frentes, TCE em Assessoria/Curso, MiniGuia separado).
  Mês da atividade = mês do `prazo` (ano < 2026 vira 1, ano > 2026 vira 12) mais os
  meses citados no `prazo_obs`. Cobertura por projeto abaixo de 30% joga o projeto
  todo para 1 a 12 (`sem_data_no_cronograma`). Linhas de projetos sem cronograma
  (FINEP, Fundraising, Glocal, Processo Seletivo, SERGIPE, alguns IU) são preservadas
  do gantt antigo pela origem `alocacao_equipe`.
  **Antes de aplicar, revalide o `gantt.py` contra o master antigo**: derivar a partir
  das Atividades antigas tem que reproduzir o `Alocacao_Gantt` antigo linha a linha.
- `Carga_Esperada` é consolidada das abas "Carga Esperada" dos cronogramas
  (CAIXA ALTA, mais granular que os nomes de projeto das Atividades, e isso é
  proposital). Refaça só os projetos cujo cronograma mudou.

## 9. Farol de percepção (drop-in)

Se vier um `Farol_Intensidade_CLEAR_2026.xlsx` **mais novo que o do GitHub**, só
substituir o arquivo no repo. Aba "Farol de Intensidade", cabeçalho na linha 4, uma
linha por pessoa, colunas Jan a Dez, escala 0 a 3 (0 sem dados, 1 tranquilo, 2 ok,
3 sobrecarga), e linha final "TOTAL EQUIPE".

## 10. Schema do master (11 abas)

- **README** atualizar "Última atualização" para hoje.
- **Atividades** `id_atividade, projeto, sub_projeto, atividade, detalhe, responsaveis,
  responsaveis_raw, n_responsaveis, prazo, prazo_obs, status, eh_entregavel, esforco,
  fonte, presencial`.
- **Resp_Atividade** uma linha por (atividade × responsável): `id_atividade, pessoa,
  projeto, prazo, status, eh_entregavel`.
- **Eventos** `id_evento, categoria, evento, tipo, projeto, data, data_fim, formato,
  local, publico_alvo, responsaveis, status, pecas_divulgacao, obs`. Vem da aba
  Eventos do cronograma de Comunicação, pulando a linha de exemplo do template.
  Status próprios: Confirmado, A confirmar, Em negociação, Realizado, Cancelado.
- **Monitorias** `id_monitoria, monitoria, data, horario, mentor, uf_grupo, turma,
  status, obs`. Vem da aba Monitorias do CPP 1.2. **A coluna de e-mail da origem
  não entra no master**, porque o repositório é público e o painel não precisa dela.
- **Pessoas** os 16 canônicos. **Projetos** lista derivada das Atividades.
- **Alocacao, Envolvimento** preservadas. **Alocacao_Gantt, Carga_Esperada** derivadas.

## 11. Checklist de validação

- [ ] Zero vazamento de "Zinho", "Pleno", "Senior", "Luiggi", "BiaS", "Samuely" em
      qualquer aba, inclusive no `responsaveis_raw`.
- [ ] `Resp_Atividade` só com canônicos e externos reais.
- [ ] `Pessoas` igual aos 16 canônicos, `id_atividade` único, zero órfão em Resp.
- [ ] `prazo` é datetime, seriais convertidos, nenhum ano fora de 2025 a 2028.
- [ ] `status` só com os 5 valores que o `app.py` conhece.
- [ ] `Alocacao` e `Envolvimento` idênticas ao master antigo. README com a data de hoje.
- [ ] ILUSOMA não aparece, EVALAC preservado, MiniGuia dentro de Bens Públicos,
      Rede PARES e IU - Formação CPP como projetos próprios.
- [ ] Diff contra o master vivo: volume por projeto, atividades novas e removidas,
      e mudanças de status, prazo e responsáveis nas linhas que existem nos dois.
      Use `collections.Counter` (multiset), não comparação de conjuntos.

## 12. Saída e deploy

Gerar o `CLEAR_Master_2026.xlsx`, entregar para download, subir no GitHub (`main`),
aguardar uns 2 minutos, Reboot se necessário. Se houve Farol novo e mais recente que
o do GitHub, subir junto.

---

## Pendências abertas (12/10/2026)

- 244 atividades sem responsável. A maior parte é vazia na origem: 53 em
  IU - Avaliação Executiva e 38 em IU - Formação CPP ("A definir", "Mentores", "—").
  Só o responsável do cronograma resolve.
- "Policy Huib" em 4 linhas de Bens Públicos: não dá pra saber se é pessoa, hub ou typo.
- "Junior person" em 3 linhas do PRiME: vaga não preenchida, não é pessoa.
- Thania está preservada como externa. Decidir se entra na equipe canônica.
- Abas `Mentores` (CPP) e `Produtos` (PRiME, TCE) sem contrapartida no painel.
