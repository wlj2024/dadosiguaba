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
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), '#f0f0f0'),
                ('GRID', (0, 0), (-1, -1), 1, '#000000')
            ]))
            elements.append(kpi_table)
            elements.append(Spacer(1, 12))

            elements.append(Paragraph("Filtros Aplicados", styles['Heading2']))
            elements.append(Paragraph(filtros_txt.replace("\n", "<br/>"), styles['Normal']))
            elements.append(Spacer(1, 12))

            elements.append(Paragraph("Gráficos", styles['Heading2']))
            fig1_img = fig_to_png(fig1)
            elements.append(Image(fig1_img, width=5*inch, height=3*inch))
            elements.append(Spacer(1, 12))
            fig2_img = fig_to_png(fig2)
            elements.append(Image(fig2_img, width=5*inch, height=3*inch))

            doc.build(elements)
            return buffer.getvalue()

        # Função de geocodificação
        def geocode_address(address):
            if 'google_api_key' not in st.session_state or st.session_state['google_api_key'] is None:
                st.error("Por favor, insira e salve uma chave de API do Google Maps nas configurações.")
                return (None, None, "Falha: Chave de API não configurada")
            geolocator = GoogleV3(api_key=st.session_state['google_api_key'])
            try:
                location = geolocator.geocode(address, timeout=10)
                if location:
                    return (location.latitude, location.longitude, "Sucesso")
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

        # Adicionando geocodificação
        with st.spinner("Geocodificando endereços..."):
            df_filtered['Address'] = df_filtered.apply(format_address, axis=1)
            geocoding_results = []
            for address in df_filtered['Address']:
                result = geocode_address(address)
                geocoding_results.append(result)
                time.sleep(1)
            df_filtered[['Latitude', 'Longitude', 'Geocoding_Status']] = pd.DataFrame(geocoding_results, index=df_filtered.index)

        # Mostrar tabela e botões de exportação
        if show_table:
            st.subheader("📄 Tabela de Empresas")
            st.dataframe(df_filtered, use_container_width=True)

            col_a, col_b, col_c = st.columns([1, 1, 1])
            with BytesIO() as buffer:
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_filtered.to_excel(writer, index=False, sheet_name="Empresas")
                    filtros_df = pd.DataFrame({"Filtros Aplicados": [filtros_txt]})
                    filtros_df.to_excel(writer, index=False, sheet_name="Filtros")
                excel_data = buffer.getvalue()
            col_a.download_button("📥 Exportar Excel", excel_data, "dados_filtrados_iguaba.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            csv_data = df_filtered.to_csv(index=False)
            col_b.download_button("📄 Exportar CSV", csv_data, "dados_filtrados_iguaba.csv", "text/csv")
            pdf_data = create_pdf()
            col_c.download_button("📑 Exportar PDF", pdf_data, "relatorio_iguaba.pdf", "application/pdf")

        # Mapa de Empresas
        st.subheader("🗺️ Mapa de Empresas")
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("Mostrar Marcadores", key="show_all_markers_btn"):
                st.session_state['show_all_markers'] = True
        with col2:
            if st.button("Limpar Marcadores", key="clear_all_markers_btn"):
                st.session_state['show_all_markers'] = False
        df_map = df_filtered.dropna(subset=['Latitude', 'Longitude'])
        if not df_map.empty and 'google_api_key' in st.session_state and st.session_state['google_api_key']:
            if st.session_state['show_all_markers']:
                markers_str = ', '.join([f"{{ lat: {row['Latitude']}, lng: {row['Longitude']}, id: '{idx}', content: '<div><strong>{row['Razao Social']}</strong><br>{row['Address']}</div>' }}" for idx, row in df_map.iterrows()])
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
                            center: {{ lat: -22.839907518453305, lng: -42.22680494297199 }},  // Centro de Iguaba Grande
                            zoom: 13,
                        }});
                        const markers = [{markers_str}];
                        const infoWindows = [];
                        if (markers.length === 0) {{
                            document.getElementById('map').innerHTML = '<p>Nenhum marcador disponível.</p>';
                            console.error('Nenhum marcador encontrado');
                            return;
                        }}
                        markers.forEach((markerData) => {{
                            const marker = new google.maps.Marker({{
                                position: {{ lat: markerData.lat, lng: markerData.lng }},
                                map: map,
                                id: markerData.id
                            }});
                            const infoWindow = new google.maps.InfoWindow({{ content: markerData.content }});
                            infoWindows.push({{ marker: marker, infoWindow: infoWindow, isOpen: false }});
                            marker.addListener('click', () => {{
                                const existing = infoWindows.find(w => w.marker.id === marker.id);
                                if (existing) {{
                                    if (existing.isOpen) {{
                                        existing.infoWindow.close();
                                        existing.isOpen = false;
                                    }} else {{
                                        existing.infoWindow.open(map, marker);
                                        existing.isOpen = true;
                                    }}
                                }}
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
                st.write("Clique em 'Mostrar Marcadores' para exibir os marcadores no mapa.")
        else:
            st.warning("Nenhum endereço pôde ser geocodificado ou a chave de API não foi configurada.")

        # Marcadores Personalizados
        st.subheader("Marcadores Personalizados")
        column = st.selectbox("Coluna", df_filtered.columns)
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            vermelho = st.selectbox("Vermelho", [None] + df_filtered[column].dropna().unique().tolist())
        with col2:
            azul = st.selectbox("Azul", [None] + df_filtered[column].dropna().unique().tolist())
        with col3:
            amarelo = st.selectbox("Amarelo", [None] + df_filtered[column].dropna().unique().tolist())
        with col4:
            verde = st.selectbox("Verde", [None] + df_filtered[column].dropna().unique().tolist())

        col5, col6 = st.columns([1, 1])
        with col5:
            if st.button("Mostrar Marcadores", key="show_custom_markers_btn"):
                st.session_state['show_custom_markers'] = True
        with col6:
            if st.button("Limpar Marcadores", key="clear_custom_markers_btn"):
                st.session_state['show_custom_markers'] = False

        if st.session_state['show_custom_markers']:
            df_map = df_filtered.dropna(subset=['Latitude', 'Longitude'])
            if not df_map.empty and 'google_api_key' in st.session_state and st.session_state['google_api_key']:
                markers = []
                for idx, row in df_map.iterrows():
                    value = row[column]
                    icon = null
                    if value == vermelho:
                        icon = "'http://maps.google.com/mapfiles/ms/icons/red-dot.png'"
                    elif value == azul:
                        icon = "'http://maps.google.com/mapfiles/ms/icons/blue-dot.png'"
                    elif value == amarelo:
                        icon = "'http://maps.google.com/mapfiles/ms/icons/yellow-dot.png'"
                    elif value == verde:
                        icon = "'http://maps.google.com/mapfiles/ms/icons/green-dot.png'"
                    markers.append(f"{{ lat: {row['Latitude']}, lng: {row['Longitude']}, id: '{idx}', content: '<div><strong>{row['Razao Social']}</strong><br>{row['Address']}</div>', icon: {icon} }}")
                markers_str = ', '.join(markers)
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
                            center: {{ lat: -22.839907518453305, lng: -42.22680494297199 }},  // Centro de Iguaba Grande
                            zoom: 13,
                        }});
                        const markers = [{markers_str}];
                        const infoWindows = [];
                        if (markers.length === 0) {{
                            document.getElementById('map').innerHTML = '<p>Nenhum marcador disponível.</p>';
                            console.error('Nenhum marcador encontrado');
                            return;
                        }}
                        markers.forEach((markerData) => {{
                            const marker = new google.maps.Marker({{
                                position: {{ lat: markerData.lat, lng: markerData.lng }},
                                map: map,
                                id: markerData.id,
                                icon: markerData.icon || null
                            }});
                            const infoWindow = new google.maps.InfoWindow({{ content: markerData.content }});
                            infoWindows.push({{ marker: marker, infoWindow: infoWindow, isOpen: false }});
                            marker.addListener('click', () => {{
                                const existing = infoWindows.find(w => w.marker.id === marker.id);
                                if (existing) {{
                                    if (existing.isOpen) {{
                                        existing.infoWindow.close();
                                        existing.isOpen = false;
                                    }} else {{
                                        existing.infoWindow.open(map, marker);
                                        existing.isOpen = true;
                                    }}
                                }}
                            }});
                        }});
                        console.log('Mapa inicializado com', markers.length, 'marcadores');
                    }}
                    window.initMap = initMap;
                </script>
                <script src="https://maps.googleapis.com/maps/api/js?key={st.session_state['google_api_key']}&callback=initMap" async defer></script>
                """
                st.components.v1.html(map_html, height=650, width=1200, scrolling=True)
                st.success(f"{len(df_map)} endereços geocodificados com marcadores personalizados.")
            else:
                st.warning("Nenhum endereço pôde ser geocodificado ou a chave de API não foi configurada.")

        # Mostrar endereços não geocodificados
        if show_failed_addresses:
            st.subheader("📍 Endereços Não Geolocalizados")
            df_failed = df_filtered[df_filtered['Latitude'].isna()]
            if not df_failed.empty:
                st.dataframe(df_failed[['Razao Social', 'Address', 'Geocoding_Status']], use_container_width=True)
            else:
                st.info("Todos os endereços foram geocodificados com sucesso ou nenhum endereço está presente nos filtros.")

# Fim do bloco principal
