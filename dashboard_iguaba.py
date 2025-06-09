import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut
import folium
from streamlit_folium import st_folium
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import plotly.io as pio
import time

# Setting up the page configuration
st.set_page_config(page_title="Dashboard Iguaba", layout="wide")
st.title("📊 Dashboard de Empresas - Iguaba Grande")

# File uploader
uploaded_file = st.file_uploader("📂 Importar planilha Excel", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Verifying required columns
    required_cols = ['Situacao Cadastral', 'Porte da Empresa', 'Optante Simples', 'Logradouro', 'Numero', 'CEP', 'Bairro', 'Municipio', 'UF']
    if not all(col in df.columns for col in required_cols):
        st.error("A planilha está faltando colunas obrigatórias: " + ", ".join([col for col in required_cols if col not in df.columns]))
    else:
        # Sidebar filters
        with st.sidebar:
            st.header("🔍 Filtros")
            situacao = st.multiselect("Situação Cadastral", df['Situacao Cadastral'].dropna().unique())
            porte = st.multiselect("Porte da Empresa", df['Porte da Empresa'].dropna().unique())
            simples = st.multiselect("Optante pelo Simples", df['Optante Simples'].dropna().unique())
            show_table = st.checkbox("📋 Mostrar tabela completa", value=True)

        # Applying filters
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

        # Function to geocode addresses
        @st.cache_data
        def geocode_address(address):
            geolocator = Nominatim(user_agent="iguaba_dashboard")
            try:
                location = geolocator.geocode(address, timeout=10)
                return (location.latitude, location.longitude) if location else (None, None)
            except (GeocoderTimedOut, Exception):
                return (None, None)

        # Preparing address for geocoding
        def format_address(row):
            parts = [
                str(row['Logradouro']) if pd.notnull(row['Logradouro']) else '',
                str(row['Numero']) if pd.notnull(row['Numero']) else '',
                str(row['Bairro']) if pd.notnull(row['Bairro']) else '',
                str(row['Municipio']) if pd.notnull(row['Municipio']) else '',
                str(row['UF']) if pd.notnull(row['UF']) else '',
                str(row['CEP']) if pd.notnull(row['CEP']) else ''
            ]
            return ', '.join([part for part in parts if part])

        # Adding geocoding to filtered dataframe
        st.subheader("🗺️ Mapa de Empresas")
        with st.spinner("Geocodificando endereços..."):
            df_filtered['Address'] = df_filtered.apply(format_address, axis=1)
            df_filtered[['Latitude', 'Longitude']] = df_filtered['Address'].apply(
                lambda x: pd.Series(geocode_address(x))
            )

        # Filtering out rows without valid coordinates
        df_map = df_filtered.dropna(subset=['Latitude', 'Longitude'])

        if not df_map.empty:
            # Creating Folium map
            m = folium.Map(location=[-22.839, -42.103], zoom_start=13)  # Centered on Iguaba Grande
            for idx, row in df_map.iterrows():
                folium.Marker(
                    [row['Latitude'], row['Longitude']],
                    popup=row['Razao Social'],
                    tooltip=row['Razao Social']
                ).add_to(m)
            st_folium(m, width=1200, height=600)
        else:
            st.warning("Nenhum endereço pôde ser geocodificado com os filtros aplicados.")

        # Gráficos
        st.subheader("📊 Gráficos")
        col4, col5 = st.columns(2)
        
        with col4:
            fig1 = px.histogram(df_filtered, x='Porte da Empresa', title="Distribuição por Porte")
            st.plotly_chart(fig1, use_container_width=True)

        with col5:
            fig2 = px.histogram(df_filtered, x='Situacao Cadastral', title="Distribuição por Situação Cadastral")
            st.plotly_chart(fig2, use_container_width=True)

        # Function to convert Plotly figure to PNG
        def fig_to_png(fig):
            img_bytes = pio.to_image(fig, format="png")
            return BytesIO(img_bytes)

        # PDF Export
        def create_pdf():
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter)
            styles = getSampleStyleSheet()
            elements = []

            # Title
            elements.append(Paragraph("Relatório de Empresas - Iguaba Grande", styles['Title']))
            elements.append(Spacer(1, 12))

            # KPIs
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
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), '#f0f0f0'),
                ('GRID', (0, 0), (-1, -1), 1, '#000000')
            ]))
            elements.append(kpi_table)
            elements.append(Spacer(1, 12))

            # Filters
            elements.append(Paragraph("Filtros Aplicados", styles['Heading2']))
            elements.append(Paragraph(filtros_txt.replace("\n", "<br/>"), styles['Normal']))
            elements.append(Spacer(1, 12))

            # Charts
            elements.append(Paragraph("Gráficos", styles['Heading2']))
            
            # Chart 1
            fig1_img = fig_to_png(fig1)
            elements.append(Image(fig1_img, width=5*inch, height=3*inch))
            elements.append(Spacer(1, 12))

            # Chart 2
            fig2_img = fig_to_png(fig2)
            elements.append(Image(fig2_img, width=5*inch, height=3*inch))

            doc.build(elements)
            return buffer.getvalue()

        # Showing table and export buttons
        if show_table:
            st.subheader("📄 Tabela de Empresas")
            st.dataframe(df_filtered, use_container_width=True)

            col_a, col_b, col_c = st.columns([1, 1, 1])

            # Export Excel
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

            # Export CSV
            csv_data = df_filtered.to_csv(index=False)
            col_b.download_button(
                label="📄 Exportar CSV",
                data=csv_data,
                file_name="dados_filtrados_iguaba.csv",
                mime="text/csv"
            )

            # Export PDF
            pdf_data = create_pdf()
            col_c.download_button(
                label="📑 Exportar PDF",
                data=pdf_data,
                file_name="relatorio_iguaba.pdf",
                mime="application/pdf"
            )
else:
    st.warning("🔁 Por favor, envie uma planilha Excel para começar.")
