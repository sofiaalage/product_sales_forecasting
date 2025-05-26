# Stock & Sales Forecasting Analysis

## Description

This Streamlit application provides an analysis of stock levels against sales forecasts. It helps users determine if current stock is sufficient to meet future demand, considering product expiration dates and customer-specific shelf-life requirements. The application simulates sequential stock consumption to provide a realistic forecast.

Users can upload an Excel file containing data on:
- Current stock on hand (including product descriptions, quantities, and expiration dates)
- Historical shipments (e.g., 2024 shipments, used to forecast for 2025)
- Customer-specific minimum shelf-life requirements

The application then processes this data to generate a detailed analysis, KPIs, and visualizations.

## Features

- **Dynamic Data Upload:** Users can upload their own Excel data files.
- **Sales Forecasting:** Forecasts future sales (e.g., for June-December 2025) based on historical data (e.g., same months in 2024).
- **Shelf-Life Consideration:** Incorporates minimum shelf-life requirements per customer.
- **Sequential Stock Consumption:** Simulates how stock is depleted over time by fulfilling orders sequentially.
- **Detailed Matrix View:** Shows a breakdown of each forecasted order, its stock status (Yes, No - Quantity, No - Validity, No - Quantity & Validity), and any missing quantities.
- **KPIs & Charts:**
    - Stock Capacity Percentage (items meeting all criteria).
    - Total Quantity Missing across all forecasts.
    - Monthly forecasted order quantities (bar chart).
    - Overall stock sufficiency by quantity (pie chart).
- **Hierarchical View:** Allows users to drill down from product level to individual customer orders.
- **Customizable UI:** Styled interface for a better user experience.

## How to Run

1.  **Clone the repository (if applicable) or ensure you have the project files.**
2.  **Install dependencies:**
    ```bash
    pip install streamlit pandas openpyxl plotly matplotlib
    ```
3.  **Navigate to the project directory:**
    ```bash
    cd path\to\your\project_directory
    ```
4.  **Run the Streamlit application:**
    ```bash
    streamlit run app.py
    ```
5.  The application will open in your web browser. Upload your Excel file when prompted.

## Dependencies

- streamlit
- pandas
- openpyxl (for reading Excel files)
- plotly
- matplotlib

## File Structure (Simplified)

```
.
├── app.py                      # Main Streamlit application script
├── Forecasting_mariana.xlsx    # Example/default data file (if not using uploader)
├── images/
│   └── logo-chiesi-footer.png  # Logo image
├── README.md                   # This file
└── ... (other project files)
```

