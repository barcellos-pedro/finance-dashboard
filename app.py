import re
import pandas as pd
import streamlit as st
import plotly.express as px

# Configuração da página do Streamlit
st.set_page_config(page_title="Dashboard de Gastos",
                   page_icon="💳", layout="wide")

# ==========================================
# Funções Utilitárias (Baseadas no seu script)
# ==========================================


def br_amount_to_float(value) -> float | None:
    if pd.isna(value):
        return None

    text = str(value).strip().replace('"', "")
    if not text:
        return None

    negative = text.startswith("-")
    if negative:
        text = text[1:].strip()

    text = text.replace(".", "").replace(",", ".")
    try:
        parsed = float(text)
        return -parsed if negative else parsed
    except ValueError:
        try:
            return float(value)
        except (TypeError, ValueError):
            return None


def excel_serial_to_date(value):
    try:
        return pd.to_datetime("1899-12-30") + pd.to_timedelta(int(float(value)), unit="D")
    except (TypeError, ValueError):
        return pd.NaT


def should_ignore_description(text: str) -> bool:
    lowered = text.lower()
    return "pagamento recebido" in lowered or "pagamento com saldo" in lowered


def suggest_cat(description) -> str:
    rules = [
        ("Uber|Metro", "Transporte"),
        ("Vivo|Google One|Spotify|Apple", "Assinaturas/Serviços"),
        ("Steam|Amazon|Shein|Samsung|Ticketmaster|Lola|Joalheria|Cosmeticos", "Compras"),
        ("Airbnb", "Viagem"),
        ("Zé Delivery|Duco|Espetus|Bella|Distribuidora De Carne|Milk Tea", "Alimentação"),
        ("Pix", "Pix/Transferência"),
        ("IOF", "Taxas"),
        ("Pagamento recebido|Pagamento Com Saldo", "Pagamento/Crédito"),
    ]

    for pattern, category in rules:
        if re.search(pattern, str(description), re.I):
            return category
    return "Outros"

# ==========================================
# Funções de Leitura de Arquivos (Uploads)
# ==========================================


def process_uploaded_files(uploaded_files) -> pd.DataFrame:
    rows = []

    for file in uploaded_files:
        filename = file.name.lower()
        file.seek(0)  # Garante que a leitura comece do início

        if filename.endswith(".csv"):
            # Processamento Nubank
            try:
                df = pd.read_csv(file, dtype=str)
                for _, row in df.iterrows():
                    title = str(row.get("title", "")).strip()
                    if should_ignore_description(title):
                        continue

                    amount = br_amount_to_float(row.get("amount"))
                    if amount is None:
                        continue

                    tipo = "Pagamento/Crédito" if amount < 0 else "Compra"
                    rows.append({
                        "Data": pd.to_datetime(str(row.get("date", "")), errors="coerce"),
                        "Descrição": title,
                        "Parcelamento": "",
                        "Valor": amount,
                        "Cartão": "Nubank",
                        "Tipo": tipo,
                    })
            except Exception as e:
                st.error(f"Erro ao ler arquivo Nubank ({file.name}): {e}")

        elif filename.endswith(".xlsx"):
            # Processamento Itaú
            if "9136" in filename:
                card_name = "Itaú Click"
            elif "1536" in filename:
                card_name = "Itaú Uniclass"
            else:
                card_name = "Itaú"

            try:
                raw = pd.read_excel(file, header=None, engine="openpyxl")
                header_idx = None
                for index in range(len(raw)):
                    values = [str(value).strip()
                              for value in raw.iloc[index].tolist()]
                    if "Data" in values and "Lançamento" in values and "Valor" in values:
                        header_idx = index
                        break

                if header_idx is not None:
                    headers = [str(value).strip() if not pd.isna(
                        value) else "" for value in raw.iloc[header_idx].tolist()]
                    data = raw.iloc[header_idx + 1:].copy()
                    data.columns = headers

                    for _, row in data.iterrows():
                        desc = row.get("Lançamento")
                        if pd.isna(desc):
                            continue

                        desc_text = str(desc).strip()
                        if should_ignore_description(desc_text):
                            continue

                        lowered = desc_text.lower()
                        if lowered.startswith("subtotal") or lowered.startswith("importante"):
                            continue

                        val = row.get("Valor")
                        amount = float(val) if not pd.isna(val) else None
                        if amount is None:
                            continue

                        dt = excel_serial_to_date(row.get("Data"))
                        tipo = "Pagamento/Crédito" if amount < 0 else "Compra"

                        rows.append({
                            "Data": dt,
                            "Descrição": desc_text,
                            "Parcelamento": "" if pd.isna(row.get("Parcelamento")) else str(row.get("Parcelamento")).strip(),
                            "Valor": amount,
                            "Cartão": card_name,
                            "Tipo": tipo,
                        })
            except Exception as e:
                st.error(f"Erro ao ler arquivo Itaú ({file.name}): {e}")

    df = pd.DataFrame(rows)

    if not df.empty:
        df = df.sort_values(["Data", "Cartão", "Descrição"],
                            na_position="last").reset_index(drop=True)
        df["Categoria"] = df["Descrição"].apply(suggest_cat)
        df["Parcelamento"] = df["Parcelamento"].fillna("")
        df["Data"] = pd.to_datetime(df["Data"]).dt.date
        return df[["Data", "Descrição", "Categoria", "Valor", "Cartão", "Parcelamento", "Tipo"]].copy()

    return pd.DataFrame(columns=["Data", "Descrição", "Categoria", "Valor", "Cartão", "Parcelamento", "Tipo"])


# ==========================================
# Interface do Streamlit
# ==========================================

st.title("📊 Controle de Gastos com Cartão")
st.markdown(
    "Faça o upload das suas faturas do Nubank (CSV) e Itaú (Excel) para visualizar o consolidado.")

# Área de Upload
uploaded_files = st.file_uploader("Selecione os arquivos", type=[
                                  'csv', 'xlsx'], accept_multiple_files=True)

if uploaded_files:
    # Processa os dados
    df_consolidado = process_uploaded_files(uploaded_files)

    if df_consolidado.empty:
        st.warning("Nenhum dado válido encontrado nos arquivos.")
    else:
        # Cálculos de Resumo
        total_compras = df_consolidado[df_consolidado["Valor"] > 0]["Valor"].sum(
        )
        total_creditos = df_consolidado[df_consolidado["Valor"] < 0]["Valor"].sum(
        )
        saldo_liquido = df_consolidado["Valor"].sum()

        st.divider()

        # KPIs (Métricas principais)
        col1, col2, col3 = st.columns(3)
        col1.metric("Total de Compras", f"R$ {total_compras:,.2f}".replace(
            ",", "X").replace(".", ",").replace("X", "."))
        col2.metric("Pagamentos/Créditos", f"R$ {total_creditos:,.2f}".replace(
            ",", "X").replace(".", ",").replace("X", "."))
        col3.metric("Saldo Líquido", f"R$ {saldo_liquido:,.2f}".replace(
            ",", "X").replace(".", ",").replace("X", "."))

        st.divider()

        # Configurando os Gráficos
        st.subheader("Análise Visual")

        # Filtra apenas os valores positivos (compras) para os gráficos fazerem sentido
        df_compras = df_consolidado[df_consolidado["Valor"] > 0]

        grafico_col1, grafico_col2 = st.columns(2)

        with grafico_col1:
            # Gráfico de Pizza: Gastos por Categoria
            if not df_compras.empty:
                fig_pie = px.pie(
                    df_compras,
                    values='Valor',
                    names='Categoria',
                    title="Gastos por Categoria",
                    # Transforma em gráfico de rosca (opcional, visualmente mais limpo)
                    hole=0.4,
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                fig_pie.update_traces(
                    textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig_pie, use_container_width=True)

        with grafico_col2:
            # Gráfico de Barras: Gastos por Cartão (Visualização Extra)
            if not df_compras.empty:
                df_cartao = df_compras.groupby(
                    'Cartão')['Valor'].sum().reset_index()
                fig_bar = px.bar(
                    df_cartao,
                    x='Cartão',
                    y='Valor',
                    title="Gastos por Cartão",
                    text_auto='.2s',
                    color='Cartão',
                    color_discrete_sequence=px.colors.qualitative.Set2
                )
                fig_bar.update_layout(showlegend=False)
                st.plotly_chart(fig_bar, use_container_width=True)

        st.divider()

        # Tabela de Dados Interativa
        st.subheader("Lançamentos Consolidados")

        # Formatando a coluna de Valor como moeda para exibição na tabela
        df_display = df_consolidado.copy()

        # Utilizando o st.dataframe para permitir ordenação e scroll
        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Valor": st.column_config.NumberColumn(
                    "Valor (R$)",
                    format="R$ %.2f"
                ),
                "Data": st.column_config.DateColumn(
                    "Data",
                    format="DD/MM/YYYY"
                )
            }
        )
else:
    st.info("Aguardando upload dos arquivos para gerar os relatórios.")
