# arquivo: dashboard_iguaba.py
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide")
st.title("📊 Dashboard de Empresas - Iguaba Grande")

uploaded_file = st.file_uploader("📥 Importe uma planilha Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="ORIGINAL")

    # Filtros
    col1, col2, col3 = st.columns(3)
    with col1:
        situacao = st.multiselect("Situação Cadastral", df['Situacao Cadastral'].dropna().unique())
    with col2:
        porte = st.multiselect("Porte da Empresa", df['Porte da Empresa'].dropna().unique())
    with col3:
        simples = st.multiselect("Optante Simples", df['Optante Simples'].dropna().unique())

    # Aplicar filtros
    df_filtered = df[
        (df['Situacao Cadastral'].isin(situacao) if situacao else True) &
        (df['Porte da Empresa'].isin(porte) if porte else True) &
        (df['Optante Simples'].isin(simples) if simples else True)
    ]

    st.markdown("## 📈 Indicadores")
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Total de Empresas", len(df_filtered))
    kpi2.metric("Ativas", df_filtered[df_filtered['Situacao Cadastral'] == "ATIVA"].shape[0])
    kpi3.metric("Com Simples", df_filtered[df_filtered['Optante Simples'] == "Sim"].shape[0])

    st.markdown("## 📊 Gráficos")
    graf1, graf2 = st.columns(2)
    with graf1:
        fig1 = px.pie(df_filtered, names="Porte da Empresa", title="Distribuição por Porte")
        st.plotly_chart(fig1, use_container_width=True)
    with graf2:
        fig2 = px.histogram(df_filtered, x="Situacao Cadastral", title="Empresas por Situação", color="Situacao Cadastral")
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("## 🧾 Tabela Detalhada")
    st.dataframe(df_filtered)
