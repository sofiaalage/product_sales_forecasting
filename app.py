import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import os
import plotly.express as px
import plotly.io as pio

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

@st.cache_data
def load_excel_data(source_file): # Modified to accept a source_file (path or uploaded file object)
    """Loads data from the Excel file and caches it."""
    stock_df = pd.read_excel(source_file, sheet_name="Stock On hand")
    shipments_df = pd.read_excel(source_file, sheet_name="2024_Shipments")
    shelf_life_df = pd.read_excel(source_file, sheet_name="shelf life")
    return stock_df, shipments_df, shelf_life_df

st.set_page_config(layout="wide", page_title="Stock & Sales Forecasting")
st.markdown(
    """
    <style>
    /* Aplica o cursor 'pointer' a todo o corpo do documento */
    body {
        cursor: pointer !important;
    }
    main{
            cursor: pointer !important;

    }
    /* Opcional: Se houver áreas específicas que ainda não mudam,
       você pode adicionar mais seletores como .st-emotion-cache-xyz
       ou div[data-testid] que englobem o conteúdo principal. */
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

# --- File Uploader Section ---
st.markdown(
    "<h3 style='font-family: Verdana, sans-serif; color: #a21a5e; font-size: 1.3em; margin-top: 20px; margin-bottom: 10px;'>Upload Your Data File</h3>",
    unsafe_allow_html=True
)
st.markdown(
    "<p style='font-family: Arial, sans-serif; color: #5d5d5d; font-weight: normal; font-size: 0.85em;'>Please upload the Excel file containing stock, shipments, and shelf life data.</p>",
    unsafe_allow_html=True
)
uploaded_file = st.file_uploader("Choose an Excel file (.xlsx or .xls)", type=["xlsx", "xls"], label_visibility="collapsed")

if uploaded_file is not None:
    try:
        stock_df, shipments_2024_df, shelf_life_df = load_excel_data(uploaded_file)
        st.success("Excel data loaded successfully!")

    except Exception as e:
        st.error(f"Failed to load data from the uploaded Excel file: {e}")
        st.stop() # Stop execution if loading fails

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

    # Ensure this part is only reached if data loading was successful
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
                    in_stock_status = "No (Quantity & Validity)" # Not enough stock, and what's left isn't valid
                    missing_qty = forecasted_qty # The whole order is missing
                    current_available_qty = 0 # Deplete remaining stock (even if not valid, it's used up attempting)

            row_result = row.copy()
            row_result['Available Stock Quantity (Initial)'] = initial_stock_info['total_available_stock'].iloc[0] if not initial_stock_info.empty else 0
            row_result['Expiration Date (Stock)'] = stock_expiration_date
            row_result['In Stock Status'] = in_stock_status
            row_result['Missing Quantity'] = missing_qty
            row_result['Remaining Stock After Order'] = current_available_qty # For tracking depletion
            sequential_analysis_results.append(row_result)

    final_analysis_df_sequential = pd.DataFrame(sequential_analysis_results)

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

        final_analysis_df_sequential['Required Expiration Date (Customer)'] = final_analysis_df_sequential.apply(
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
        
        final_matrix_sequential = final_analysis_df_sequential[display_columns].copy()
        
        final_matrix_sequential['Forecast Ship Date'] = final_matrix_sequential['Forecast Ship Date'].dt.strftime('%Y-%m-%d')
        final_matrix_sequential['Expiration Date (Stock)'] = final_matrix_sequential['Expiration Date (Stock)'].dt.strftime('%Y-%m-%d').replace({pd.NaT: 'N/A'})
        final_matrix_sequential['Required Expiration Date (Customer)'] = final_matrix_sequential['Required Expiration Date (Customer)'].dt.strftime('%Y-%m-%d')

        st.dataframe(final_matrix_sequential, use_container_width=True)


        
        total_forecasted_qty = final_analysis_df_sequential['Forecasted Qty'].sum()
        total_missing_qty = final_analysis_df_sequential['Missing Quantity'].sum()
        
        items_fully_covered = final_analysis_df_sequential[final_analysis_df_sequential['In Stock Status'] == 'Yes'].shape[0]
        total_forecasted_items = final_analysis_df_sequential.shape[0]
        
        percentage_capacity = 0
        if total_forecasted_items > 0:
            percentage_capacity = (items_fully_covered / total_forecasted_items) * 100
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
        col1, col2= st.columns(2)
        with col1:
            st.metric(
                label="Stock Capacity (Items Meeting All Criteria - Sequential)",
                value=f"{percentage_capacity:.2f}%"
            )
        with col2:
            st.metric(
                label="Total Quantity Missing Across All Forecasts (Sequential)",
                value=f"{total_missing_qty:,.0f}"
            )

        st.markdown(
            """
            <p style='font-family: Arial, sans-serif; color: #5d5d5d; font-weight: normal; font-size: 0.7em;'>
            *Stock Capacity represents the percentage of forecasted product-customer-month combinations
            that can be fully met by current stock, considering both available quantity and the required shelf-life,
            after accounting for stock depletion by earlier orders.
            </p>
            """,
            unsafe_allow_html=True
        )

        st.markdown("---")    



        col1, col2= st.columns(2)
        with col1:

            st.markdown(
                "<h4 style='font-family: Verdana, sans-serif; color: #a21a5e; font-size: 1.2em;'>Monthly Forecasted Orders (June - Dec 2025)</h4>",
                unsafe_allow_html=True
            )

            monthly_forecast = final_analysis_df_sequential.groupby(final_analysis_df_sequential['Forecast Ship Date'].dt.to_period('M'))['Forecasted Qty'].sum().reset_index()
            monthly_forecast['Month'] = monthly_forecast['Forecast Ship Date'].dt.strftime('%b %Y')

            fig_px_bar = px.bar(
                monthly_forecast,
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


            st.markdown(
                """
                <style>
                /* Altera o cursor para 'pointer' ao passar sobre a área do gráfico Plotly */
                .plotly-graph-div {
                    cursor: pointer !important;
                }
                </style>
                """,
                unsafe_allow_html=True
            )


            st.plotly_chart(fig_px_bar, use_container_width=True)




        with col2:
            # Título do Gráfico de Suficiência: mesma fonte e cor vinho
# Título do Gráfico de Suficiência: mesma fonte e cor vinho
            st.markdown(
                "<h4 style='font-family: Verdana, sans-serif; color: #a21a5e; font-size: 1.2em;'>Overall Stock Sufficiency by Quantity (Based on Sequential Analysis)</h4>",
                unsafe_allow_html=True
            )
            total_covered_qty = total_forecasted_qty - total_missing_qty


            if total_forecasted_qty > 0:
                # Dados para o gráfico de pizza
                pie_data = pd.DataFrame({
                    'Category': ['Quantity Covered', 'Quantity Missing'],
                    'Quantity': [total_covered_qty, total_missing_qty]
                })

                # Cores para as fatias da pizza
                colors = ['#a21a5e', '#c7c8c9']  # Verde para Covered, Vinho para Missing

                # Criar o gráfico de pizza (donut) com Plotly Express
                fig_pie = px.pie(
                    pie_data,
                    values='Quantity',
                    names='Category',
                    color='Category',
                    color_discrete_map={
                        'Quantity Covered': '#a21a5e',
                        'Quantity Missing': '#c7c8c9'
                    }
                )

                # Ordenar as fatias (opcional, mas recomendado para melhor legibilidade)
                fig_pie.update_traces(
                    sort=False,  # Mantém a ordem original dos dados
                    textposition='inside',
                    textinfo='percent+label',
                )

                # Ajustar o layout do gráfico
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
                st.info("No forecasted quantities to display for sufficiency analysis.")



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

        # Group by Product Description for a hierarchical view using expanders
        for product in final_matrix_sequential['Item Description'].unique():
            product_df = final_matrix_sequential[final_matrix_sequential['Item Description'] == product]
            
            # Calculate product-level summaries
            product_total_forecasted = product_df['Forecasted Qty'].sum()
            product_total_missing = product_df['Missing Quantity'].sum()

            with st.expander(f"**Product:** {product} (Total Forecasted: {product_total_forecasted:,.0f} | Missing: {product_total_missing:,.0f})"):
                st.markdown(f"**Customer Details for {product}:**")
                # Iterate through customers and display their data directly within the product expander
                for customer in product_df['Ship To Customer (Bill To)'].unique():
                    customer_df = product_df[product_df['Ship To Customer (Bill To)'] == customer]
                    
                    customer_shelf_life = customer_df['Min Shelf-Life (Months)'].iloc[0] if not customer_df.empty else 'N/A'
                    customer_total_forecasted = customer_df['Forecasted Qty'].sum()
                    customer_total_missing = customer_df['Missing Quantity'].sum()

                    # Using a smaller header or just bold text for customer details
                    st.markdown(f"**Customer:** {customer} | Req. Shelf-life: {customer_shelf_life} months (Forecasted: {customer_total_forecasted:,.0f} | Missing: {customer_total_missing:,.0f})")
                    st.dataframe(customer_df[[
                        'Forecast Ship Date',
                        'Forecasted Qty',
                        'Available Stock Quantity (Initial)',
                        'Remaining Stock After Order',
                        'Expiration Date (Stock)',
                        'Required Expiration Date (Customer)',
                        'In Stock Status',
                        'Missing Quantity'
                    ]].reset_index(drop=True), use_container_width=True)

        st.markdown("---")
else:
    st.info("Please upload an Excel file to proceed with the analysis.")
    st.stop() # Stop if no file is uploaded
