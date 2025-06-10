import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from geopy.geocoders import GoogleV3
from geopy.exc import GeocoderTimedOut, GeocoderUnavailable
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import plotly.io as pio
import time

# Configuração da página
st.set_page_config(page_title="Dashboard Iguaba", layout="wide")
st.title("📊 Dashboard de Empresas - Iguaba Grande")

# Inicializar estados
if 'google_api_key' not in st.session_state:
    st.session_state['google_api_key'] = "AIzaSyAxkuSyZfTZc9cD2UjUJFV0rXqLkf1yFzQ"
if 'show_key_input' not in st.session_state:
    st.session_state['show_key_input'] = False
if 'show_all_markers' not in st.session_state:
    st.session_state['show_all_markers'] = False
if 'show_custom_markers' not in st.session_state:
    st.session_state['show_custom_markers'] = False

# Campo para chave de API na barra lateral
with st.sidebar:
    st.header("⚙️ Configurações da API")
    if 'google_api_key' in st.session_state and not st.session_state['show_key_input']:
        masked_key = "**********..." + st.session_state['google_api_key'][-4:]
        st.write("Chave da API salva:", masked_key)
        if st.button("Deletar Chave", key="delete_key"):
            del st.session_state['google_api_key']
            st.session_state['show_key_input'] = True
    if st.session_state['show_key_input'] or 'google_api_key' not in st.session_state:
        api_key = st.text_input("Insira a Chave da API do Google Maps", key="api_key_input")
        if st.button("Salvar Chave", key="save_key"):
            st.session_state['google_api_key'] = api_key
            st.session_state['show_key_input'] = False
    st.write("A chave acima é de testes, se desejar usar sua própria chave, clique em Deletar chave, cole sua chave e salve. Para voltar a usar a chave de teste basta atualizar essa página.")

# Upload do arquivo
uploaded_file = st.file_uploader("📂 Importar planilha Excel", type=["xlsx"], key="file_uploader")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Verificar colunas obrigatórias
    required_cols = ['Situacao Cadastral', 'Porte da Empresa', 'Optante Simples', 'Logradouro', 'Numero', 'CEP', 'Bairro', 'Municipio', 'UF']
    if not all(col in df.columns for col in required_cols):
        st.error("A planilha está faltando colunas obrigatórias: " + ", ".join([col for col in required_cols if col not in df.columns]))
    else:
        # Aplicar filtros
        with st.sidebar:
            st.header("🔍 Filtros")
            situacao = st.multiselect("Situação Cadastral", df['Situacao Cadastral'].dropna().unique())
            porte = st.multiselect("Porte da Empresa", df['Porte da Empresa'].dropna().unique())
            simples = st.multiselect("Optante pelo Simples", df['Optante Simples'].dropna().unique())
            show_table = st.checkbox("📋 Mostrar tabela completa", value=True)
            show_failed_addresses = st.checkbox("📍 Mostrar endereços não geocodificados", value=False)

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
        total_empresas = len(df_filtered)
        empresas_ativas = df_filtered[df_filtered['Situacao Cadastral'] == 'ATIVA'].shape[0]
        optantes_simples = df_filtered[df_filtered['Optante Simples'] == 'Sim'].shape[0]
        col1.metric("Total de Empresas", total_empresas)
        col2.metric("Empresas Ativas", empresas_ativas)
        col3.metric("Optantes do Simples", optantes_simples)

        # Gráficos
        st.subheader("📊 Gráficos")
        col4, col5 = st.columns(2)
        
        with col4:
            chart_type_porte = st.selectbox(
                "Tipo de gráfico para Porte",
                ["barras", "pizza", "linha", "área"],
                index=1
            )
            if chart_type_porte == "barras":
                fig1 = px.histogram(df_filtered, x='Porte da Empresa', title="Distribuição por Porte")
            elif chart_type_porte == "pizza":
                fig1 = px.pie(df_filtered, names='Porte da Empresa', title="Distribuição por Porte")
            elif chart_type_porte == "linha":
                fig1 = px.line(df_filtered, x='Porte da Empresa', title="Distribuição por Porte")
            elif chart_type_porte == "área":
                fig1 = px.area(df_filtered, x='Porte da Empresa', title="Distribuição por Porte")
            st.plotly_chart(fig1, use_container_width=True)

        with col5:
            chart_type_situacao = st.selectbox(
                "Tipo de gráfico para Situação Cadastral",
                ["barras", "pizza", "linha", "área"],
                index=0
            )
            if chart_type_situacao == "barras":
                fig2 = px.histogram(df_filtered, x='Situacao Cadastral', title="Distribuição por Situação Cadastral")
            elif chart_type_situacao == "pizza":
                fig2 = px.pie(df_filtered, names='Situacao Cadastral', title="Distribuição por Situação Cadastral")
            elif chart_type_situacao == "linha":
                fig2 = px.line(df_filtered, x='Situacao Cadastral', title="Distribuição por Situação Cadastral")
            elif chart_type_situacao == "área":
                fig2 = px.area(df_filtered, x='Situacao Cadastral', title="Distribuição por Situação Cadastral")
            st.plotly_chart(fig2, use_container_width=True)

        # Função para converter gráfico Plotly em PNG
        def fig_to_png(fig):
            img_bytes = pio.to_image(fig, format="png")
            return BytesIO(img_bytes)

        # Exportação para PDF
        def create_pdf():
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter)
            styles = getSampleStyleSheet()
            elements = []

            elements.append(Paragraph("Relatório de Empresas - Iguaba Grande", styles['Title']))
            elements.append(Spacer(1, 12))

            elements.append(Paragraph("KPIs", styles['Heading2']))
            kpi_data = [
                ["Métrica", "Valor"],
                ["Total de Empresas", str(total_empresas)],
                ["Empresas Ativas", str(empresas_ativas)],
                ["Optantes do Simples", str(optantes_simples)]
            ]
            kpi_table = Table(kpi_data)
            kpi_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), '#d0d0d0'),
                ('TEXTCOLOR', (0, 0), (-1, 0), '#000000'),
                ('ALIGN
