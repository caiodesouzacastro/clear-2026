# CLEAR 2026 — Dashboard de Organização e Alocação

App Streamlit pra visualizar o portfólio de projetos, alocação de pesquisadores e linha de chegada (entregas críticas) do FGV CLEAR em 2026.

## Estrutura

```
clear-2026/
├── app.py                     # App Streamlit
├── CLEAR_Master_2026.xlsx     # Fonte única de dados (editável)
├── requirements.txt
└── README.md
```

A planilha mestra é a **única fonte de dados**. Editou a planilha → o app reflete.

## Rodar localmente

```bash
# 1. Clonar repositório
git clone <url-do-repo>
cd clear-2026

# 2. Criar ambiente (opcional mas recomendado)
python -m venv .venv
source .venv/bin/activate    # Linux/Mac
# .venv\Scripts\activate     # Windows

# 3. Instalar dependências
pip install -r requirements.txt

# 4. Rodar
streamlit run app.py
```

Abre em `http://localhost:8501`.

## Hospedar grátis (Streamlit Community Cloud)

1. Faça push do repositório pro GitHub (público ou privado)
2. Entre em https://share.streamlit.io
3. "New app" → selecione o repo → arquivo principal: `app.py`
4. Deploy

A cada push novo no GitHub, o app atualiza automaticamente.

## As 3 visões

### 📊 Portfólio
Visão executiva do portfólio inteiro: heatmap de densidade de atividades por projeto × mês, timeline de atividades, lista de próximos entregáveis. Pergunta que responde: *o portfólio está saudável?*

### 👤 Pesquisador
Pra cada pessoa CLEAR: gráfico de ocupação mensal (% da capacidade alocada), timeline de atividades sob responsabilidade, tabela detalhada. Pergunta: *quem está sobrecarregado?*

### 🏁 Linha de Chegada
Calendário de entregas críticas agrupado por urgência (atrasados → 7 dias → 30 dias → depois). Pergunta: *o que preciso despachar essa semana?*

## Filtros (sidebar)

- **Período**: próximas 4 semanas, 3 meses, resto do ano ou tudo
- **Projetos**: subset do portfólio
- **Status**: Concluído, Em Andamento, Não Iniciado, Atrasado, Reunião

## Manutenção dos dados

A aba `Atividades` da planilha mestra tem as colunas:

| Campo | Descrição |
|---|---|
| `id_atividade` | ID único (ACT0001...) |
| `projeto` | Projeto pai |
| `sub_projeto` | Subdivisão (módulo, etapa) |
| `atividade` | Descrição |
| `responsaveis` | Pessoas CLEAR separadas por `;` (use os nomes canônicos da aba `Pessoas`) |
| `prazo` | Data limite |
| `status` | `Concluído` / `Em Andamento` / `Não Iniciado` / `Atrasado` / `Reunião` |
| `eh_entregavel` | `TRUE` para entregáveis críticos (aparecem na Linha de Chegada) |

**Para adicionar uma atividade nova:** continue a numeração de `id_atividade`, preencha os campos, salve.

## 18 pessoas CLEAR (nomes canônicos)

Bia B, Bia S, Caio, Carol, Cecilia, Fabrícia, Julia, Junior, Lorena, Luan, Luiggi, Lycia, Michel, PMO, Pleno, Samu, Senior, Zinho
