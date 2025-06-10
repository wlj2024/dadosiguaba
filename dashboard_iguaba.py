import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderUnavailable
import folium
from streamlit_folium import st_folium
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import plotly.io as pio
import time
import folium.plugins
import random  # Para variação opcional nas coordenadas

# Configuração da página
st.set_page_config(page_title="Dashboard Iguaba", layout="wide")
st.title("📊 Dashboard de Empresas - Iguaba Grande")

# Upload do arquivo
uploaded_file = st.file_uploader("📂 Importar planilha Excel", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Verificar colunas obrigatórias
    required_cols = ['Situacao Cadastral', 'Porte da Empresa', 'Optante Simples', 'Logradouro', 'Numero', 'CEP', 'Bairro', 'Municipio', 'UF']
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
            show_failed_addresses = st.checkbox("📍 Mostrar endereços não geocodificados", value=False)

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
        total_empresas = len(df_filtered)
        empresas_ativas = df_filtered[df_filtered['Situacao Cadastral'] == 'ATIVA'].shape[0]
        optantes_simples = df_filtered[df_filtered['Optante Simples'] == 'Sim'].shape[0]
        col1.metric("Total de Empresas", total_empresas)
        col2.metric("Empresas Ativas", empresas_ativas)
        col3.metric("Optantes do Simples", optantes_simples)

        # Função de geocodificação com múltiplas tentativas
        @st.cache_data
        def geocode_address(address, municipio="IGUABA GRANDE, RJ"):
            geolocator = Nominatim(user_agent="iguaba_dashboard")
            try:
                # Tentativa 1: Endereço completo
                location = geolocator.geocode(address, timeout=10)
                if location:
                    return (location.latitude, location.longitude, "Sucesso")
                # Tentativa 2: Rua + Município + UF
                rua_municipio = ', '.join([part for part in address.split(', ') if part.lower() not in ['sn', municipio.lower()] and not part.isdigit()][:2] + [municipio])
                location = geolocator.geocode(rua_municipio, timeout=10)
                if location:
                    return (location.latitude, location.longitude, "Sucesso com rua e município")
                # Tentativa 3: Rua + CEP
                rua_cep = ', '.join([part for part in address.split(', ') if part.replace('-', '').isdigit() or not part.isdigit()][:2])
                location = geolocator.geocode(rua_cep, timeout=10)
                if location:
                    return (location.latitude, location.longitude, "Sucesso com rua e CEP")
                # Tentativa 4: Apenas Município (fallback)
                fallback_address = municipio
                location = geolocator.geocode(fallback_address, timeout=10)
                if location:
                    return (location.latitude, location.longitude, "Sucesso com fallback (município)")
                return (None, None, f"Falha: Nenhum resultado para {address}")
            except GeocoderTimedOut:
                return (None, None, f"Falha: Timeout para {address}")
            except GeocoderUnavailable:
                return (None, None, f"Falha: Serviço indisponível para {address}")
            except Exception as e:
                return (None, None, f"Falha: {str(e)}")

        # Formatando endereço
        def format_address(row):
            parts = [
                str(row['Logradouro']) if pd.notnull(row['Logradouro']) else '',
                str(row['Numero']) if pd.notnull(row['Numero']) and str(row['Numero']).lower() != 'sn' else '',
                str(row['Bairro']) if pd.notnull(row['Bairro']) else '',
                str(row['Municipio']) if pd.notnull(row['Municipio']) else '',
                str(row['UF']) if pd.notnull(row['UF']) else '',
                str(row['CEP']) if pd.notnull(row['CEP']) else ''
            ]
            return ', '.join([part for part in parts if part])

        # Adicionando geocodificação ao dataframe filtrado
        st.subheader("🗺️ Mapa de Empresas")
        with st.spinner("Geocodificando endereços..."):
            df_filtered['Address'] = df_filtered.apply(format_address, axis=1)
            geocoding_results = []
            for address in df_filtered['Address']:
                result = geocode_address(address)
                # Adicionar pequena variação se usar fallback
                if result[2] == "Sucesso com fallback (município)":
                    lat = result[0] + random.uniform(-0.005, 0.005)  # Variação de ~500m
                    lon = result[1] + random.uniform(-0.005, 0.005)
                    geocoding_results.append((lat, lon, result[2]))
                else:
                    geocoding_results.append(result)
                time.sleep(1)  # Atraso para respeitar limites do Nominatim
            df_filtered[['Latitude', 'Longitude', 'Geocoding_Status']] = pd.DataFrame(geocoding_results, index=df_filtered.index)

        # Filtrando empresas com coordenadas válidas
        df_map = df_filtered.dropna(subset=['Latitude', 'Longitude'])

        if not df_map.empty:
            # Criando mapa Folium com MarkerCluster
            m = folium.Map(location=[-22.839, -42.103], zoom_start=13)  # Centrado em Iguaba Grande
            marker_cluster = folium.plugins.MarkerCluster().add_to(m)
            for idx, row in df_map.iterrows():
                folium.Marker(
                    [row['Latitude'], row['Longitude']],
                    popup=f"{row['Razao Social']}<br>{row['Address']}",
                    tooltip=row['Razao Social']
                ).add_to(marker_cluster)
            st_folium(m, width=1200, height=600)
            st.success(f"{len(df_map)} endereços geocodificados com sucesso.")
        else:
            st.warning("Nenhum endereço pôde ser geocodificado com os filtros aplicados.")

        # Mostrar endereços não geocodificados (se selecionado)
        if show_failed_addresses:
            st.subheader("📍 Endereços Não Geocodificados")
            df_failed = df_filtered[df_filtered['Latitude'].isna()]
            if not df_failed.empty:
                st.dataframe(df_failed[['Razao Social', 'Address', 'Geocoding_Status']], use_container_width=True)
            else:
                st.info("Todos os endereços foram geocodificados com sucesso ou nenhum endereço está presente nos filtros.")

        # Gráficos
        st.subheader("📊 Gráficos")
        col4, col5 = st.columns(2)
        
        with col4:
            fig1 = px.histogram(df_filtered, x='Porte da Empresa', title="Distribuição por Porte")
            st.plotly_chart(fig1, use_container_width=True)

        with col5:
            fig2 = px.histogram(df_filtered, x='Situacao Cadastral', title="Distribuição por Situação Cadastral")
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

            # Título
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

            # Filtros
            elements.append(Paragraph("Filtros Aplicados", styles['Heading2']))
            elements.append(Paragraph(filtros_txt.replace("\n", "<br/>"), styles['Normal']))
            elements.append(Spacer(1, 12))

            # Gráficos
            elements.append(Paragraph("Gráficos", styles['Heading2']))
            
            # Gráfico 1
            fig1_img = fig_to_png(fig1)
            elements.append(Image(fig1_img, width=5*inch, height=3*inch))
            elements.append(Spacer(1, 12))

            # Gráfico 2
            fig2_img = fig_to_png(fig2)
            elements.append(Image(fig2_img, width=5*inch, height=3*inch))

            doc.build(elements)
            return buffer.getvalue()

        # Mostrar tabela e botões de exportação
        if show_table:
            st.subheader("📄 Tabela de Empresas")
            st.dataframe(df_filtered, use_container_width=True)

            col_a, col_b, col_c = st.columns([1, 1, 1])

            # Exportar Excel
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

            # Exportar CSV
            csv_data = df_filtered.to_csv(index=False)
            col_b.download_button(
                label="📄 Exportar CSV",
                data=csv_data,
                file_name="dados_filtrados_iguaba.csv",
                mime="text/csv"
            )

            # Exportar PDF
            pdf_data = create_pdf()
            col_c.download_button(
                label="📑 Exportar PDF",
                data=pdf_data,
                file_name="relatorio_iguaba.pdf",
                mime="application/pdf"
            )
else:
    st.warning("🔁 Por favor, envie uma planilha Excel para começar.")
