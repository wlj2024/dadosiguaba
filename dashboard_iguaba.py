import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Dashboard Iguaba", layout="wide")
st.title("📊 Dashboard de Empresas - Iguaba Grande")

uploaded_file = st.file_uploader("📂 Importar planilha Excel", type=["xlsx"])

def gerar_excel_com_filtros(df, filtros):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Dados Filtrados")
        filtros_df = pd.DataFrame({"Filtro": list(filtros.keys()), "Valor": list(filtros.values())})
        filtros_df.to_excel(writer, index=False, sheet_name="Filtros Aplicados")
    return output.getvalue()

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    required_cols = ['Situacao Cadastral', 'Porte da Empresa', 'Optante Simples']
    if not all(col in df.columns for col in required_cols):
        st.error("A planilha está faltando colunas obrigatórias: " + ", ".join([col for col in required_cols if col not in df.columns]))
    else:
        with st.sidebar:
            st.header("🔍 Filtros")
            situacao = st.multiselect("Situação Cadastral", df['Situacao Cadastral'].dropna().unique())
            porte = st.multiselect("Porte da Empresa", df['Porte da Empresa'].dropna().unique())
            simples = st.multiselect("Optante pelo Simples", df['Optante Simples'].dropna().unique())
            show_table = st.checkbox("📋 Mostrar tabela completa", value=True)

        df_filtered = df.copy()
        if situacao:
            df_filtered = df_filtered[df_filtered['Situacao Cadastral'].isin(situacao)]
        if porte:
            df_filtered = df_filtered[df_filtered['Porte da Empresa'].isin(porte)]
        if simples:
            df_filtered = df_filtered[df_filtered['Optante Simples'].isin(simples)]

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

        # Tabela + exportações
        if show_table:
            st.subheader("📄 Tabela de Empresas")
            st.dataframe(df_filtered, use_container_width=True)

            col_export1, col_export2 = st.columns(2)

            with col_export1:
                st.download_button(
                    label="📥 Exportar para Excel",
                    data=gerar_excel_com_filtros(df_filtered, {
                        "Situação Cadastral": ', '.join(situacao) if situacao else "Todos",
                        "Porte da Empresa": ', '.join(porte) if porte else "Todos",
                        "Optante Simples": ', '.join(simples) if simples else "Todos"
                    }),
                    file_name="dados_filtrados_iguaba.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with col_export2:
                st.download_button(
                    label="📄 Exportar para CSV",
                    data=df_filtered.to_csv(index=False).encode('utf-8'),
                    file_name="dados_filtrados_iguaba.csv",
                    mime="text/csv"
                )

            # Exibe os filtros aplicados como log
            st.markdown("### 🧾 Filtros Aplicados")
            st.json({
                "Situação Cadastral": situacao if situacao else "Todos",
                "Porte da Empresa": porte if porte else "Todos",
                "Optante Simples": simples if simples else "Todos"
            })

else:
    st.warning("🔁 Por favor, envie uma planilha Excel para começar.")
