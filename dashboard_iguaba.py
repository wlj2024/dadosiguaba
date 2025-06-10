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

# Campo para chave de API na barra lateral
with st.sidebar:
    st.header("⚙️ Configurações da API")
    if 'google_api_key' not in st.session_state or st.session_state['google_api_key'] is None:
        api_key = st.text_input("Insira a Chave da API do Google Maps", "")
        if st.button("Salvar Chave"):
            st.session_state['google_api_key'] = api_key
            st.rerun()
    else:
        masked_key = "**********..." + st.session_state['google_api_key'][-4:]
        st.write("Chave da API salva:", masked_key)
        if st.button("Deletar Chave"):
            del st.session_state['google_api_key']
            st.rerun()

    st.header("🔍 Filtros")
    situacao = st.multiselect("Situação Cadastral", df['Situacao Cadastral'].dropna().unique() if 'df' in locals() else [])
    porte = st.multiselect("Porte da Empresa", df['Porte da Empresa'].dropna().unique() if 'df' in locals() else [])
    simples = st.multiselect("Optante pelo Simples", df['Optante Simples'].dropna().unique() if 'df' in locals() else [])
    show_table = st.checkbox("📋 Mostrar tabela completa", value=True)
    show_failed_addresses = st.checkbox("📍 Mostrar endereços não geocodificados", value=False)

# Upload do arquivo
uploaded_file = st.file_uploader("📂 Importar planilha Excel", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Verificar colunas obrigatórias
    required_cols = ['Situacao Cadastral', 'Porte da Empresa', 'Optante Simples', 'Logradouro', 'Numero', 'CEP', 'Bairro', 'Municipio', 'UF']
    if not all(col in df.columns for col in required_cols):
        st.error("A planilha está faltando colunas obrigatórias: " + ", ".join([col for col in required_cols if col not in df.columns]))
    else:
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

        # Função de geocodificação com Google Maps API
        def geocode_address(address):
            if 'google_api_key' not in st.session_state or st.session_state['google_api_key'] is None:
                st.error("Por favor, insira e salve uma chave de API do Google Maps nas configurações.")
                return (None, None, "Falha: Chave de API não configurada")
            geolocator = GoogleV3(api_key=st.session_state['google_api_key'])
            try:
                # Tentativa com endereço completo
                location = geolocator.geocode(address, timeout=10)
                if location:
                    return (location.latitude, location.longitude, "Sucesso")
                # Tentativa simplificada (sem número)
                simplified_address = ', '.join([part for part in address.split(', ') if not part.isdigit()])
                location = geolocator.geocode(simplified_address, timeout=10)
                if location:
                    return (location.latitude, location.longitude, "Sucesso sem número")
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
        with st.spinner("Geocodificando endereços..."):
            df_filtered['Address'] = df_filtered.apply(format_address, axis=1)
            geocoding_results = []
            for address in df_filtered['Address']:
                result = geocode_address(address)
                geocoding_results.append(result)
                time.sleep(1)  # Atraso para respeitar limites da API
            df_filtered[['Latitude', 'Longitude', 'Geocoding_Status']] = pd.DataFrame(geocoding_results, index=df_filtered.index)

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

        # Filtrando empresas com coordenadas válidas para o mapa
        df_map = df_filtered.dropna(subset=['Latitude', 'Longitude'])

        # Mapa de Empresas
        st.subheader("🗺️ Mapa de Empresas")
        if not df_map.empty and 'google_api_key' in st.session_state and st.session_state['google_api_key']:
            markers_str = ', '.join([f"{{ lat: {row['Latitude']}, lng: {row['Longitude']}, title: '{row['Razao Social']}<br>{row['Address']}' }}" for _, row in df_map.iterrows()])
            map_html = f"""
            <div id="map" style="height: 600px; width: 1200px; border: 1px solid #ccc;"></div>
            <script>
                function initMap() {{
                    if (!window.google || !window.google.maps) {{
                        document.getElementById('map').innerHTML = '<p style="color:red;">Erro: A API do Google Maps não carregou. Verifique a chave de API ou a conexão.</p>';
                        console.error('Google Maps API não carregada');
                        return;
                    }}
                    const map = new google.maps.Map(document.getElementById("map"), {{
                        center: {{ lat: -22.839, lng: -42.103 }},
                        zoom: 13,
                    }});
                    const markers = [{markers_str}];
                    if (markers.length === 0) {{
                        document.getElementById('map').innerHTML = '<p>Nenhum marcador disponível.</p>';
                        console.error('Nenhum marcador encontrado');
                        return;
                    }}
                    markers.forEach((marker) => {{
                        new google.maps.Marker({{
                            position: {{ lat: marker.lat, lng: marker.lng }},
                            map: map,
                            title: marker.title
                        }});
                    }});
                    console.log('Mapa inicializado com', markers.length, 'marcadores');
                }}
                window.initMap = initMap;
            </script>
            <script src="https://maps.googleapis.com/maps/api/js?key={st.session_state['google_api_key']}&callback=initMap" async defer></script>
            """
            st.components.v1.html(map_html, height=650, width=1200, scrolling=True)
            st.success(f"{len(df_map)} endereços geocodificados com sucesso.")
        else:
            st.warning("Nenhum endereço pôde ser geocodificado ou a chave de API não foi configurada.")

        # Mostrar endereços não geocodificados (se selecionado)
        if show_failed_addresses:
            st.subheader("📍 Endereços Não Geolocalizados")
            df_failed = df_filtered[df_filtered['Latitude'].isna()]
            if not df_failed.empty:
                st.dataframe(df_failed[['Razao Social', 'Address', 'Geocoding_Status']], use_container_width=True)
            else:
                st.info("Todos os endereços foram geocodificados com sucesso ou nenhum endereço está presente nos filtros.")

# Fim do bloco principal
