# 📊 Controle de Gastos com Cartão — Dashboard Streamlit

Este repositório contém uma aplicação interativa desenvolvida em **Streamlit** para consolidação e análise visual de faturas de cartão de crédito. A aplicação substitui e estende o script original baseado em `openpyxl`, permitindo carregar faturas dinamicamente via interface web em vez de ler arquivos locais fixos.

## 🚀 Funcionalidades

- **Upload Dinâmico de Arquivos:** Suporte para múltiplos arquivos simultâneos (faturas `.csv` e `.xlsx`.
- **Métricas em Tempo Real (KPIs):** Visualização instantânea do Total de Compras, Total de Pagamentos/Créditos e Saldo Líquido.
- **Gráficos Interativos (Plotly):**
  - **Gastos por Categoria:** Gráfico de rosca detalhando a distribuição percentual das compras.
  - **Gastos por Cartão:** Gráfico de barras comparando o volume de gastos entre as faturas enviadas.
- **Tabela de Lançamentos Consolidados:** Visualização tabular interativa com suporte a ordenação, busca e paginação nativa do Streamlit, com formatação monetária (R$) e de datas adequada.
- **Categorização Automática:** Motor de inferência baseado em Expressões Regulares (RegEx) para sugerir categorias com base na descrição dos lançamentos.

## 🛠️ Tecnologías Utilizadas

- **Python 3.10+**
- **Streamlit** (Interface de usuário e interatividade)
- **Pandas** (Tratamento, filtragem e consolidação dos dados)
- **Plotly Express** (Gráficos dinâmicos e interativos)
- **Openpyxl** (Engine de leitura dos arquivos Excel)
