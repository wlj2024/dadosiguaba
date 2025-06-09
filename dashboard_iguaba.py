import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO

st.set_page_config(page_title="Dashboard Iguaba", layout="wide")
st.title("📊 Dashboard de Empresas - Iguaba Grande")

uploaded_file = st.file_uploader("📂 Importar planilha Excel", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Verifica se as colunas necessárias existem
    required_cols = ['Situacao Cadastral', 'Porte da Empresa', 'Optante Simples']
    if not all(col in df.columns for col in required_cols):
        st.error("A planilha está faltando colunas obrigatórias: " + ", ".join([col for col in required_cols if col not in df.columns]))
    else:
        # Filtros na barra lateral
        with st.sidebar:
            st.header("🔍 Filtros")
            situacao = st.multiselect("Situação Cadastral", df['Situacao Cadastral'].dropna().unique())
            porte = st.multiselect("Porte da Empresa", df['Porte da Empresa'].dropna().unique())
            simples = st.multiselect("Optante pelo Simples", df['Optante Simples'].dropna().unique())
            show_table = st.checkbox("📋 Mostrar tabela completa", value=True)

        # Aplicar filtros
        df_filtered = df.copy()
        filtros_aplicados = []
        if situacao:
            df_filtered = df_filtered[df_filtered['Situacao Cadastral'].isin(situacao)]
            filtros_aplicados.append(f"Situação: {', '.join(situacao)}")
        if porte:
            df_filtered = df_filtered[df_filtered['Porte da Empresa'].isin(porte)]
            filtros_aplicados.append(f"Porte: {', '.join(porte)}")
        if simples:
            df_filtered = df_filtered[df_filtered['Optante Simples'].isin(simples)]
            filtros_aplicados.append(f"Simples: {', '.join(simples)}")

        filtros_txt = "\n".join(filtros_aplicados) if filtros_aplicados else "Nenhum filtro aplicado."

        # KPIs
        st.subheader("📈 KPIs")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total de Empresas", len(df_filtered))
        col2.metric("Empresas Ativas", df_filtered[df_filtered['Situacao Cadastral'] == 'ATIVA'].shape[0])
        col3.metric("Optantes do Simples", df_filtered[df_filtered['Optante Simples'] == 'SIM'].shape[0])

        # Gráficos
        st.subheader("📊 Gráficos")
        fig1 = px.histogram(df_filtered, x='Porte da Empresa', title="Distribuição por Porte")
        st.plotly_chart(fig1, use_container_width=True)

        fig2 = px.histogram(df_filtered, x='Situacao Cadastral', title="Distribuição por Situação Cadastral")
        st.plotly_chart(fig2, use_container_width=True)

        # Mostrar tabela e botões de exportação
        if show_table:
            st.subheader("📄 Tabela de Empresas")
            st.dataframe(df_filtered, use_container_width=True)

            col_a, col_b = st.columns([1, 1])

            # Exportar para Excel
            with BytesIO() as buffer:
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_filtered.to_excel(writer, index=False, sheet_name="Empresas")
                    filtros_df = pd.DataFrame({"Filtros Aplicados": [filtros_txt]})
                    filtros_df.to_excel(writer, index=False, sheet_name="Filtros")
                excel_data = buffer.getvalue()

            col_a.download_button(
                label="📥 Exportar Excel",
                data=excel_data,
                file_name="dados_filtrados_iguaba.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Exportar para CSV
            csv_data = df_filtered.to_csv(index=False)
            col_b.download_button(
                label="📄 Exportar CSV",
                data=csv_data,
                file_name="dados_filtrados_iguaba.csv",
                mime="text/csv"
            )
else:
    st.warning("🔁 Por favor, envie uma planilha Excel para começar.")
