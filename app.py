import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# --- Helper Functions ---
def parse_shelf_life(text):
    """
    Parses the 'Minimum Shelf-life' text from the Excel sheet into an integer
    representing months. Defaults to 6 months if the text is ambiguous or empty.
    """
    if pd.isna(text) or not isinstance(text, str):
        return 6  # Default to 6 months if NaN or not a string
    text_lower = text.lower()
    if "12 months" in text_lower or "1 year" in text_lower or "not less than 12 months" in text_lower:
        return 12
    elif "6 months" in text_lower:
        return 6
    elif "3 months" in text_lower:
        return 3
    else:
        return 6  

# Initialize session state variables if not already present
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
    st.session_state.final_analysis_df = None
    st.session_state.available_months = []
    st.session_state.available_customers = []
    st.session_state.available_medicines = []

@st.cache_data
def load_excel_data(source_file): # Modified to accept a source_file (path or uploaded file object)
    """Loads data from the Excel file and caches it."""
    stock_df = pd.read_excel(source_file, sheet_name="Stock On hand")
    shipments_df = pd.read_excel(source_file, sheet_name="2024_Shipments")
    shelf_life_df = pd.read_excel(source_file, sheet_name="shelf life")
    return stock_df, shipments_df, shelf_life_df

st.set_page_config(layout="wide", page_title="Stock & Sales Forecasting")

# --- Main application logic: Conditional display of uploader or tabs ---
if not st.session_state.data_loaded:
    # --- File Uploader Section ---
    st.markdown(
        "<h3 style='font-family: Verdana, sans-serif; color: #a21a5e; font-size: 1.3em; margin-top: 20px; margin-bottom: 10px;'>Upload Your Data File</h3>",
        unsafe_allow_html=True
    )
    uploaded_file = st.file_uploader("Choose an Excel file (.xlsx or .xls)", type=["xlsx", "xls"], label_visibility="collapsed")

    if uploaded_file is not None:
        try:
            stock_df, shipments_2024_df, shelf_life_df = load_excel_data(uploaded_file)
            # Removed: st.success("Excel data loaded successfully!") - will be hidden by rerun

            # --- Start of Data Processing ---
            stock_df['Expiration Date'] = pd.to_datetime(stock_df['Expiration Date'], errors='coerce')
            stock_df = stock_df.rename(columns={'Description': 'Product Description', 'Available To Reserve': 'Available Stock Quantity'})

            consolidated_stock = stock_df.groupby('Product Description').agg(
                total_available_stock=('Available Stock Quantity', 'sum'),
                earliest_expiration_date=('Expiration Date', 'min')
            ).reset_index()
            consolidated_stock = consolidated_stock.rename(columns={'Product Description': 'Item Description'})

            shelf_life_df['Min Shelf-Life (Months)'] = shelf_life_df['Minimum Shelf-life (reported on customer PO)'].apply(parse_shelf_life)
            shelf_life_df = shelf_life_df.rename(columns={'Customer Name': 'Ship To Customer (Bill To)'})
            shelf_life_df = shelf_life_df.drop(columns=['Minimum Shelf-life (reported on customer PO)'])

            shipments_2024_df['Ship Date'] = pd.to_datetime(shipments_2024_df['Ship Date'], errors='coerce')
            shipments_2024_df = shipments_2024_df.dropna(subset=['Ship Date'])

            shipments_2024_filtered = shipments_2024_df[
                (shipments_2024_df['Ship Date'].dt.month >= 6) &
                (shipments_2024_df['Ship Date'].dt.month <= 12)
            ].copy()

            forecast_2025_df = shipments_2024_filtered.groupby([
                'Item Description',
                'Ship To Customer (Bill To)',
                shipments_2024_filtered['Ship Date'].dt.to_period('M')
            ])['Qty'].sum().reset_index()

            forecast_2025_df['Forecast Ship Date'] = forecast_2025_df['Ship Date'].dt.start_time.apply(lambda x: x.replace(year=2025))
            forecast_2025_df = forecast_2025_df.rename(columns={'Qty': 'Forecasted Qty'})

            forecasted_orders_with_shelf_life = pd.merge(
                forecast_2025_df,
                shelf_life_df,
                on='Ship To Customer (Bill To)',
                how='left'
            )
            forecasted_orders_with_shelf_life['Min Shelf-Life (Months)'] = forecasted_orders_with_shelf_life['Min Shelf-Life (Months)'].fillna(6)

            sequential_analysis_results = []
            for product_name, product_group_df in forecasted_orders_with_shelf_life.groupby('Item Description'):
                initial_stock_info = consolidated_stock[consolidated_stock['Item Description'] == product_name]
                current_available_qty = initial_stock_info['total_available_stock'].iloc[0] if not initial_stock_info.empty else 0
                stock_expiration_date = initial_stock_info['earliest_expiration_date'].iloc[0] if not initial_stock_info.empty else pd.NaT
                product_group_df_sorted = product_group_df.sort_values(by='Forecast Ship Date')

                for index, row in product_group_df_sorted.iterrows():
                    forecasted_qty = row['Forecasted Qty']
                    forecast_ship_date = row['Forecast Ship Date']
                    min_shelf_life_months = row['Min Shelf-Life (Months)']
                    required_expiration_date = forecast_ship_date + pd.DateOffset(months=min_shelf_life_months)
                    in_stock_status = "No"
                    missing_qty = forecasted_qty
                    is_stock_valid_for_order = pd.notna(stock_expiration_date) and stock_expiration_date >= required_expiration_date

                    if current_available_qty >= forecasted_qty:
                        if is_stock_valid_for_order:
                            in_stock_status = "Yes"
                            missing_qty = 0
                            current_available_qty -= forecasted_qty
                        else:
                            in_stock_status = "No (Validity)"
                    else:
                        if is_stock_valid_for_order:
                            in_stock_status = "No (Quantity)"
                            missing_qty = forecasted_qty - current_available_qty
                            current_available_qty = 0
                        else:
                            in_stock_status = "No (Quantity & Validity)"
                            missing_qty = forecasted_qty
                            current_available_qty = 0
                    
                    row_result = row.copy()
                    row_result['Available Stock Quantity (Initial)'] = initial_stock_info['total_available_stock'].iloc[0] if not initial_stock_info.empty else 0
                    row_result['Expiration Date (Stock)'] = stock_expiration_date
                    row_result['In Stock Status'] = in_stock_status
                    row_result['Missing Quantity'] = missing_qty
                    row_result['Remaining Stock After Order'] = current_available_qty
                    sequential_analysis_results.append(row_result)
            
            final_analysis_df_sequential = pd.DataFrame(sequential_analysis_results)
            # --- End of Data Processing ---

            # Store processed data in session state
            st.session_state.final_analysis_df = final_analysis_df_sequential
            if not final_analysis_df_sequential.empty:
                st.session_state.available_months = sorted(final_analysis_df_sequential['Forecast Ship Date'].dt.strftime('%Y-%m').unique())
                st.session_state.available_customers = sorted(final_analysis_df_sequential['Ship To Customer (Bill To)'].unique())
                st.session_state.available_medicines = sorted(final_analysis_df_sequential['Item Description'].unique())
            else:
                st.session_state.available_months = []
                st.session_state.available_customers = []
                st.session_state.available_medicines = []
            
            st.session_state.data_loaded = True
            st.rerun()

        except Exception as e:
            st.error(f"Failed to load data from the uploaded Excel file: {e}")
            st.session_state.data_loaded = False
            st.session_state.final_analysis_df = None
            st.session_state.available_months = []
            st.session_state.available_customers = []
            st.session_state.available_medicines = []
            st.stop()
else: # This means st.session_state.data_loaded is True. Display the tabs.
    # Retrieve data from session state
    final_analysis_df_sequential = st.session_state.final_analysis_df
    available_months = st.session_state.available_months
    available_customers = st.session_state.available_customers
    available_medicines = st.session_state.available_medicines

    if final_analysis_df_sequential is not None:
        # --- Tab definitions and content ---
        st.markdown(
            """
            <style>
            /* Estilo do texto dentro das abas */
            .stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p {
                font-size: 0.9rem; /* Diminuir a fonte do texto das abas */
                padding: 0px 10px; /* Ajusta o padding horizontal interno para 'encolher' a barra, se necessário */
            }

            /* Cor de fundo da aba selecionada (ativa) e a "barra" que simula a linha inferior */
            .stTabs [data-baseweb="tab-list"] button[aria-selected="true"] {
                background-color: #a21a5e; /* Sua cor roxa para a aba ativa */
                color: white; /* Cor do texto da aba selecionada */
                border-radius: 5px 5px 0 0; /* Bordas arredondadas apenas no topo */
                border-bottom: 3px solid #a21a5e; /* Cria uma "barra" mais grossa na parte inferior da aba ativa */
                margin-bottom: -3px; /* Compensa a borda para que não empurre o conteúdo */
            }

            /* Cor de fundo das abas não selecionadas (inativas) */
            .stTabs [data-baseweb="tab-list"] button {
                background-color: #6c6c6c; /* Um cinza mais escuro para as abas inativas */
                color: white; /* Cor do texto das abas inativas para contraste */
                border-radius: 5px 5px 0 0; /* Bordas arredondadas apenas no topo */
                border-bottom: 3px solid #6c6c6c; /* Cria uma "barra" para as abas inativas também */
            }

            /* Remover a linha inferior padrão do Streamlit nas abas */
            .stTabs [data-baseweb="tab-list"] {
                border-bottom: none !important; /* Garante que não haja linha na parte inferior da lista de abas */
            }

            /* Remover a linha inferior que o Streamlit pode adicionar quando a aba está ativa */
            .stTabs [data-baseweb="tab-list"] button[aria-selected="true"]::after {
                content: none; /* Remove completamente o elemento ::after (a linha/barra padrão) */
            }

            /* Ajustar o espaçamento entre as abas, se necessário */
            .stTabs [data-baseweb="tab-list"] {
                gap: 5px; /* Adiciona um pequeno espaçamento entre as abas */
            }
            </style>
            """,
            unsafe_allow_html=True
        )
        col_logo, col_title = st.columns([0.2, 0.8]) 

        with col_logo:

            st.image("images/logo-chiesi-footer.png", width=250)



        with col_title:

            st.markdown(
                """
                <div style=height:15px; ></div>
                <div style="display: flex; flex-direction: column; justify-content: flex-end; height: 100%;">
                    <div style="flex-grow: 1;"></div> <h2 style='font-family: Verdana, sans-serif; color: #2c2c2c; font-size: 1.5em; margin-bottom: 0;'>
                        Stock & Sales Forecasting Analysis
                    </h2>
                </div>
                """,
                unsafe_allow_html=True
            )


        text_content = """
        This application helps you assess if your current stock is sufficient to meet your
        future sales forecasts. It takes into account both the quantity available and the
        product's expiration date relative to customer-specific shelf-life requirements with sequential stock consumption simulation!
        """

        st.markdown(
            f"<p style='font-family: Arial, sans-serif; color: #5d5d5d; font-weight: normal; font-size: 0.85em;'>{text_content}</p>",
            unsafe_allow_html=True
        )


        tab_matrix, tab_kpis_charts, tab_hierarchical = st.tabs(["Detailed Matrix", "KPIs & Charts", "Hierarchical View"])



        with tab_matrix:
            st.markdown(
                "<h3 style='font-family: Verdana, sans-serif; color: #a21a5e; font-size: 1.3em;'>Detailed Stock Analysis Matrix (June - Dec 2025 Forecast - Sequential)</h3>",
                unsafe_allow_html=True
            )
            st.markdown(
                """
                <p style='font-family: Arial, sans-serif; color: #5d5d5d; font-weight: normal; font-size: 0.85em;'>
                This matrix provides a detailed breakdown of forecasted orders, now reflecting
                the sequential depletion of stock over time. It indicates whether each order
                can be fully met by current stock based on quantity and required shelf-life.
                </p>
                """,
                unsafe_allow_html=True
            )

            # --- Filters for Detailed Matrix ---
            col_filter1_matrix, col_filter2_matrix, col_filter3_matrix = st.columns(3)
            with col_filter1_matrix:
                selected_month_matrix = st.selectbox("Month (YYYY-MM):", options=['All'] + available_months, key='month_matrix')
            with col_filter2_matrix:
                selected_customer_matrix = st.selectbox("Customer:", options=['All'] + available_customers, key='customer_matrix')
            with col_filter3_matrix:
                selected_medicine_matrix = st.selectbox("Item:", options=['All'] + available_medicines, key='medicine_matrix')

            # Filter data for the matrix
            filtered_data_matrix = final_analysis_df_sequential.copy()
            if selected_month_matrix != 'All':
                filtered_data_matrix = filtered_data_matrix[filtered_data_matrix['Forecast Ship Date'].dt.strftime('%Y-%m') == selected_month_matrix]
            if selected_customer_matrix != 'All':
                filtered_data_matrix = filtered_data_matrix[filtered_data_matrix['Ship To Customer (Bill To)'] == selected_customer_matrix]
            if selected_medicine_matrix != 'All':
                filtered_data_matrix = filtered_data_matrix[filtered_data_matrix['Item Description'] == selected_medicine_matrix]


            if not filtered_data_matrix.empty:
                filtered_data_matrix['Required Expiration Date (Customer)'] = filtered_data_matrix.apply(
                    lambda row: row['Forecast Ship Date'] + pd.DateOffset(months=row['Min Shelf-Life (Months)']), axis=1
                )

                display_columns = [
                    'Item Description',
                    'Forecast Ship Date',
                    'Forecasted Qty',
                    'Ship To Customer (Bill To)',
                    'Available Stock Quantity (Initial)', # Initial stock for reference
                    'Remaining Stock After Order',      # Stock remaining after this order
                    'Expiration Date (Stock)',          # Earliest stock expiry for the product
                    'Min Shelf-Life (Months)',          # Customer requirement in months
                    'Required Expiration Date (Customer)', # Calculated minimum date
                    'In Stock Status',
                    'Missing Quantity'
                ]
                
                final_matrix_display = filtered_data_matrix[display_columns].copy()
                
                final_matrix_display['Forecast Ship Date'] = final_matrix_display['Forecast Ship Date'].dt.strftime('%Y-%m-%d')
                final_matrix_display['Expiration Date (Stock)'] = final_matrix_display['Expiration Date (Stock)'].dt.strftime('%Y-%m-%d').replace({pd.NaT: 'N/A'})
                final_matrix_display['Required Expiration Date (Customer)'] = final_matrix_display['Required Expiration Date (Customer)'].dt.strftime('%Y-%m-%d')

                st.dataframe(final_matrix_display, use_container_width=True)
            else:
                st.info("No data to display for the selected filters in the Detailed Matrix.")


            
            total_forecasted_qty = final_analysis_df_sequential['Forecasted Qty'].sum() # This remains global for context or should be filtered too?
            total_missing_qty_overall = final_analysis_df_sequential['Missing Quantity'].sum() # This remains global for context
            
            # KPIs below the matrix could also be filtered or show overall. For now, they are overall.
            # If they need to be filtered, use filtered_data_matrix for these calculations.

            st.markdown(
                """
                <style>
                /* Estilo geral para os containers das métricas (os "cards") */
                div[data-testid="stMetric"] {
                    background-color: white; /* Fundo branco para o card */
                    border-radius: 10px; /* Bordas arredondadas */
                    padding: 20px; /* Espaçamento interno */
                    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); /* Sombra suave para efeito de card */
                    margin-bottom: 20px; /* Espaçamento entre os cards e o conteúdo abaixo */
                }

                /* Estilo para o rótulo (label) da métrica */
                div[data-testid="stMetric"] label p {
                    font-family: Arial, sans-serif; /* Fonte do subtítulo da página principal */
                    color: #5d5d5d ; /* Cor cinza escuro para o rótulo */
                    font-weight: bold; /* Sem negrito */
                    font-size: 1em; /* Tamanho da fonte do rótulo */
                }

                /* Estilo para o valor (value) da métrica */
                div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
                    font-family: Verdana, sans-serif; /* Fonte do título da página principal */
                    color: #a21a5e; /* Cor vinho para o valor */
                    font-size: 2.5em; /* Aumenta o tamanho do valor para destaque */
                    font-weight: bold; /* Deixa o valor em negrito */
                }

                /* Estilo para o delta (se houver um) na métrica */
                div[data-testid="stMetricDelta"] {
                    font-family: Arial, sans-serif; /* Fonte consistente */
                    color: #5d5d5d; /* Cor cinza escuro */
                    font-size: 1em; /* Tamanho padrão */
                }
                </style>
                """,
                unsafe_allow_html=True
            )
        with tab_kpis_charts:
            # --- Filters for KPIs & Charts ---
            col_filter1_kpi, col_filter2_kpi, col_filter3_kpi = st.columns(3)
            with col_filter1_kpi:
                selected_month_kpi = st.selectbox("Month (YYYY-MM):", options=['All'] + available_months, key='month_kpi')
            with col_filter2_kpi:
                selected_customer_kpi = st.selectbox("Customer:", options=['All'] + available_customers, key='customer_kpi')
            with col_filter3_kpi:
                selected_medicine_kpi = st.selectbox("Item:", options=['All'] + available_medicines, key='medicine_kpi')

            # Filter data for KPIs and Charts
            filtered_data_kpis = final_analysis_df_sequential.copy()
            if selected_month_kpi != 'All':
                filtered_data_kpis = filtered_data_kpis[filtered_data_kpis['Forecast Ship Date'].dt.strftime('%Y-%m') == selected_month_kpi]
            if selected_customer_kpi != 'All':
                filtered_data_kpis = filtered_data_kpis[filtered_data_kpis['Ship To Customer (Bill To)'] == selected_customer_kpi]
            if selected_medicine_kpi != 'All':
                filtered_data_kpis = filtered_data_kpis[filtered_data_kpis['Item Description'] == selected_medicine_kpi]

            if not filtered_data_kpis.empty:
                kpi_total_forecasted_qty = filtered_data_kpis['Forecasted Qty'].sum()
                kpi_total_missing_qty = filtered_data_kpis['Missing Quantity'].sum()
                
                kpi_items_fully_covered = filtered_data_kpis[filtered_data_kpis['In Stock Status'] == 'Yes'].shape[0]
                kpi_total_forecasted_items = filtered_data_kpis.shape[0]
                
                kpi_percentage_capacity = 0
                if kpi_total_forecasted_items > 0:
                    kpi_percentage_capacity = (kpi_items_fully_covered / kpi_total_forecasted_items) * 100

                col1_kpi_display, col2_kpi_display = st.columns(2)
                with col1_kpi_display:
                    st.metric(
                        label="Stock Capacity (Items Meeting All Criteria - Filtered)",
                        value=f"{kpi_percentage_capacity:.2f}%"
                    )
                with col2_kpi_display:
                    st.metric(
                        label="Total Quantity Missing (Filtered)",
                        value=f"{kpi_total_missing_qty:,.0f}"
                    )
            else:
                st.info("No data for KPIs based on selected filters.")
                kpi_percentage_capacity = 0 # Default values if no data
                kpi_total_missing_qty = 0


            st.markdown(
                """
                <style>
                /* Estilo geral para os containers das métricas (os "cards") */
                div[data-testid="stMetric"] {
                    background-color: white; /* Fundo branco para o card */
                    border-radius: 10px; /* Bordas arredondadas */
                    padding: 20px; /* Espaçamento interno */
                    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); /* Sombra suave para efeito de card */
                    margin-bottom: 20px; /* Espaçamento entre os cards e o conteúdo abaixo */
                }

                /* Estilo para o rótulo (label) da métrica */
                div[data-testid="stMetric"] label p {
                    font-family: Arial, sans-serif; /* Fonte do subtítulo da página principal */
                    color: #5d5d5d ; /* Cor cinza escuro para o rótulo */
                    font-weight: bold; /* Sem negrito */
                    font-size: 1em; /* Tamanho da fonte do rótulo */
                }

                /* Estilo para o valor (value) da métrica */
                div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
                    font-family: Verdana, sans-serif; /* Fonte do título da página principal */
                    color: #a21a5e; /* Cor vinho para o valor */
                    font-size: 2.5em; /* Aumenta o tamanho do valor para destaque */
                    font-weight: bold; /* Deixa o valor em negrito */
                }

                /* Estilo para o delta (se houver um) na métrica */
                div[data-testid="stMetricDelta"] {
                    font-family: Arial, sans-serif; /* Fonte consistente */
                    color: #5d5d5d; /* Cor cinza escuro */
                    font-size: 1em; /* Tamanho padrão */
                }
                </style>
                """,
                unsafe_allow_html=True
            )
            col1_charts, col2_charts = st.columns(2) # Renamed to avoid conflict
            with col1_charts:
                st.markdown(
                    "<h4 style='font-family: Verdana, sans-serif; color: #a21a5e; font-size: 1.2em;'>Monthly Forecasted Orders (Filtered)</h4>",
                    unsafe_allow_html=True
                )
                if not filtered_data_kpis.empty:
                    monthly_forecast_chart_data = filtered_data_kpis.groupby(filtered_data_kpis['Forecast Ship Date'].dt.to_period('M'))['Forecasted Qty'].sum().reset_index()
                    monthly_forecast_chart_data['Month'] = monthly_forecast_chart_data['Forecast Ship Date'].dt.strftime('%b %Y')

                    fig_px_bar = px.bar(
                        monthly_forecast_chart_data,
                        x='Month',
                        y='Forecasted Qty'
                    )
                    fig_px_bar.update_traces(marker_color='#c7c8c9')
                    fig_px_bar.update_layout(
                        hoverlabel=dict(
                            bgcolor="#a21a5e", 
                            font_size=12,
                            font_color="#ffffff",
                            font_family="Arial, sans-serif" 
                        ),
                        plot_bgcolor='rgba(0,0,0,0)', 
                        paper_bgcolor='rgba(0,0,0,0)', 
                    )
                    st.plotly_chart(fig_px_bar, use_container_width=True)
                else:
                    st.info("No monthly forecast data to display for the selected filters.")

            with col2_charts: # Moved Pie Chart to the second column
                st.markdown(
                    "<h4 style='font-family: Verdana, sans-serif; color: #a21a5e; font-size: 1.2em;'>Overall Stock Sufficiency by Quantity (Filtered)</h4>",
                    unsafe_allow_html=True
                )
                
                if not filtered_data_kpis.empty and kpi_total_forecasted_qty > 0:
                    pie_total_covered_qty = kpi_total_forecasted_qty - kpi_total_missing_qty
                    pie_data_chart = pd.DataFrame({
                        'Category': ['Quantity Covered', 'Quantity Missing'],
                        'Quantity': [pie_total_covered_qty, kpi_total_missing_qty]
                    })
                    fig_pie = px.pie(
                        pie_data_chart,
                        values='Quantity',
                        names='Category',
                        color='Category',
                        color_discrete_map={
                            'Quantity Covered': '#a21a5e',
                            'Quantity Missing': '#c7c8c9'
                        }
                    )
                    fig_pie.update_traces(
                        sort=False,
                        textposition='inside',
                        textinfo='percent+label',
                    )
                    fig_pie.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)',
                        hoverlabel=dict(
                            bgcolor="#a21a5e",
                            font_size=12,
                            font_color="white",
                            font_family="Verdana, Geneva, sans-serif"
                        ),
                        showlegend=True,
                        legend_title_text='Status',
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=1.02,
                            xanchor="right",
                            x=1
                        )
                    )
                    st.plotly_chart(fig_pie, use_container_width=True)
                else:
                    st.info("No stock sufficiency data to display for the selected filters.")



        with tab_hierarchical:
            st.markdown(
                "<h3 style='font-family: Verdana, sans-serif; color: #a21a5e; font-size: 1.3em;'>Hierarchical View of Analysis Results (Sequential)</h3>",
                unsafe_allow_html=True
            )
            st.markdown(
                """
                <p style='font-family: Arial, sans-serif; color: #5d5d5d; font-weight: normal; font-size: 0.85em;'>
                This section provides a drill-down view of the analysis, allowing you to
                explore results from product level down to individual customer orders and their stock status,
                <b>reflecting the sequential stock consumption.</b>
                </p>
                """,
                unsafe_allow_html=True
            )

            # --- Filters for Hierarchical View ---
            col_filter1_hier, col_filter2_hier, col_filter3_hier = st.columns(3)
            with col_filter1_hier:
                selected_month_hier = st.selectbox("Month (YYYY-MM):", options=['All'] + available_months, key='month_hier')
            with col_filter2_hier:
                selected_customer_hier = st.selectbox("Customer:", options=['All'] + available_customers, key='customer_hier')
            with col_filter3_hier:
                selected_medicine_hier = st.selectbox("Item:", options=['All'] + available_medicines, key='medicine_hier')

            # Filter data for the hierarchical view
            filtered_data_hierarchical = final_analysis_df_sequential.copy()
            if selected_month_hier != 'All':
                filtered_data_hierarchical = filtered_data_hierarchical[filtered_data_hierarchical['Forecast Ship Date'].dt.strftime('%Y-%m') == selected_month_hier]
            if selected_customer_hier != 'All':
                filtered_data_hierarchical = filtered_data_hierarchical[filtered_data_hierarchical['Ship To Customer (Bill To)'] == selected_customer_hier]
            if selected_medicine_hier != 'All':
                filtered_data_hierarchical = filtered_data_hierarchical[filtered_data_hierarchical['Item Description'] == selected_medicine_hier]

            if not filtered_data_hierarchical.empty:
                # Prepare data for display (similar to matrix tab, but used within expanders)
                # Ensure 'Required Expiration Date (Customer)' is calculated for the filtered data
                filtered_data_hierarchical['Required Expiration Date (Customer)'] = filtered_data_hierarchical.apply(
                    lambda row: row['Forecast Ship Date'] + pd.DateOffset(months=row['Min Shelf-Life (Months)']), axis=1
                )
                # Format dates for display
                display_df_hierarchical = filtered_data_hierarchical.copy()
                display_df_hierarchical['Forecast Ship Date'] = display_df_hierarchical['Forecast Ship Date'].dt.strftime('%Y-%m-%d')
                display_df_hierarchical['Expiration Date (Stock)'] = display_df_hierarchical['Expiration Date (Stock)'].dt.strftime('%Y-%m-%d').replace({pd.NaT: 'N/A'})
                display_df_hierarchical['Required Expiration Date (Customer)'] = display_df_hierarchical['Required Expiration Date (Customer)'].dt.strftime('%Y-%m-%d')

                for product in display_df_hierarchical['Item Description'].unique():
                    product_df_hier = display_df_hierarchical[display_df_hierarchical['Item Description'] == product]
                    
                    product_total_forecasted = product_df_hier['Forecasted Qty'].sum() # Use original Qty for sum
                    product_total_missing = product_df_hier['Missing Quantity'].sum()

                    with st.expander(f"**Product:** {product} (Total Forecasted: {product_total_forecasted:,.0f} | Missing: {product_total_missing:,.0f})"):
                        st.markdown(f"**Customer Details for {product}:**")
                        for customer in product_df_hier['Ship To Customer (Bill To)'].unique():
                            customer_df_hier = product_df_hier[product_df_hier['Ship To Customer (Bill To)'] == customer]
                            
                            customer_shelf_life = customer_df_hier['Min Shelf-Life (Months)'].iloc[0] if not customer_df_hier.empty else 'N/A'
                            customer_total_forecasted = customer_df_hier['Forecasted Qty'].sum()
                            customer_total_missing = customer_df_hier['Missing Quantity'].sum()

                            st.markdown(f"**Customer:** {customer} | Req. Shelf-life: {customer_shelf_life} months (Forecasted: {customer_total_forecasted:,.0f} | Missing: {customer_total_missing:,.0f})")
                            st.dataframe(customer_df_hier[[
                                'Forecast Ship Date',
                                'Forecasted Qty',
                                'Available Stock Quantity (Initial)',
                                'Remaining Stock After Order',
                                'Expiration Date (Stock)',
                                'Required Expiration Date (Customer)',
                                'In Stock Status',
                                'Missing Quantity'
                            ]].reset_index(drop=True), use_container_width=True)
            else:
                st.info("No data to display for the selected filters in the Hierarchical View.")
