# CLEAR 2026 — Dashboard de Organização e Alocação

## 🔗 Acesse o painel

**👉 [https://clear-2026-ahhtunccmwdbxx5nmfxpbb.streamlit.app](https://clear-2026-ahhtunccmwdbxx5nmfxpbb.streamlit.app)**

App Streamlit para visualizar o portfólio de projetos, alocação de pesquisadores e próximas entregas do FGV CLEAR em 2026.

---

## Estrutura do repositório

```
clear-2026/
├── app.py                     # App Streamlit
├── CLEAR_Master_2026.xlsx     # Fonte única de dados (editável)
├── config.toml                # Tema visual
├── requirements.txt
└── README.md
```

A planilha mestra é a **única fonte de dados**. Editou a planilha → o app reflete em 1-2 minutos.

## As 4 abas

### 📊 Visão Geral
Linha do tempo (Gantt) por projeto, marcadores de entregáveis e KPIs. Pergunta: *como está distribuído o ano?*

### 📅 Calendário
Grade mensal navegável com as entregas do mês. Pergunta: *o que acontece este mês?*

### 👥 Equipe
Mapa de carga — pessoas × meses, intensidade pelo número de projetos simultâneos. Pergunta: *quem está sobrecarregado?*

### 👤 Pesquisador
A fundo numa pessoa: capacidade, projetos simultâneos, atividades atribuídas. Pergunta: *qual é a carga desta pessoa?*

## Filtros (sidebar)

- **Mostrar**: só entregáveis críticos ou todas as atividades
- **Projetos**: subset do portfólio

## Como atualizar os dados

1. Cada projeto preenche seu cronograma (modelo padrão `Modelo_Cronograma_CLEAR_2026.xlsx`)
2. Uma pessoa junta todos os cronogramas no Claude usando `PROMPT_gerar_planilha.md` e gera o novo `CLEAR_Master_2026.xlsx`
3. Sobe o arquivo no GitHub substituindo o atual — o painel atualiza sozinho

Detalhes no arquivo `COMO_ATUALIZAR.md`.

## 16 pessoas CLEAR (nomes canônicos)

Bia B, Bia S, Caio, Carol, Cecilia, Fabrícia, Julia, Junior, Lorena, Luan, Luigi, Lycia, Michel, Pleno, Samu, Senior

## Rodar localmente (dev)

```bash
git clone https://github.com/caiodesouzacastro/clear-2026
cd clear-2026
pip install -r requirements.txt
streamlit run app.py
```

Abre em `http://localhost:8501`.
