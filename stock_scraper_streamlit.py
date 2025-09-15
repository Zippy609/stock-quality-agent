# stock_scraper_streamlit.py
import streamlit as st
import pandas as pd
import plotly.express as px
from extract_data_from_screener_site import scrape_stock  # your scraper program
import warnings

# Suppress non-critical warnings
warnings.filterwarnings("ignore")

SECTION_NAMES = {
    'quarters': 'Quarterly P&L',
    'profit-loss': 'Yearly P&L',
    'peers': 'Peers',
    'balance-sheet': 'Balance Sheet',
    'cash-flow': 'Cash Flow',
    'ratios': 'Ratios',
    'shareholding': 'Shareholding'
}

st.set_page_config(page_title="Stock Screener Data", layout="wide")

st.title("📊 Stock Screener Data")

# Stock input
stock_symbol = st.text_input("Enter stock symbol (e.g., TCS)").upper()
load_button = st.button("Load Data")

if load_button:
    if not stock_symbol:
        st.warning("Please enter a stock symbol.")
    else:
        with st.spinner(f"Fetching data for {stock_symbol}..."):
            try:
                status_message, section_tables = scrape_stock(stock_symbol)
            except Exception as e:
                st.error(f"Error scraping stock: {e}")
                section_tables = {}
                status_message = "Error"

        st.success(status_message)

        calculated_metrics = []
        opm_chart = None

        # Display each section
        for section_id, tables in section_tables.items():
            friendly_name = SECTION_NAMES.get(section_id, section_id)
            if not tables:
                continue

            for i, df in enumerate(tables, start=1):
                if df.empty:
                    continue
                df = df.copy()
                st.subheader(f"{friendly_name} - Table {i}")
                st.dataframe(df.fillna(''))

                # Calculated metrics and QoQ chart for Quarterly P&L
                if section_id == 'quarters' and 'Parameters' in df.columns:
                    try:
                        sales_row = df[df['Parameters'].str.contains("Sales", case=False, na=False)]
                        op_row = df[df['Parameters'].str.contains("Operating Profit", case=False, na=False)]
                        quarter_cols = df.columns[1:]
                        if not sales_row.empty and not op_row.empty:
                            sales_values = pd.to_numeric(sales_row.iloc[0][quarter_cols], errors='coerce')
                            op_values = pd.to_numeric(op_row.iloc[0][quarter_cols], errors='coerce')
                            opm_values = (op_values / sales_values * 100).round(2).fillna(0)

                            # Add to Calculated Metrics
                            calc_df = pd.DataFrame({
                                'Metric': ['Operating Profit % of Sales'],
                                **{col: [val] for col, val in zip(quarter_cols, opm_values)}
                            })
                            calculated_metrics.append(calc_df)

                            # Plot chart
                            chart_df = pd.DataFrame({
                                'Quarter': quarter_cols,
                                'Operating Margin (%)': opm_values.values
                            })
                            opm_chart = px.bar(
                                chart_df,
                                x='Quarter',
                                y='Operating Margin (%)',
                                text='Operating Margin (%)',
                                color='Operating Margin (%)',
                                color_continuous_scale='Blues',
                                title='Operating Profit % of Sales (QoQ)'
                            )
                            opm_chart.update_traces(textposition='auto')
                    except Exception as e:
                        st.warning(f"Error calculating QoQ Operating Margin: {e}")

        # Show Calculated Metrics
        if calculated_metrics:
            st.subheader("📈 Calculated Metrics")
            calc_full_df = pd.concat(calculated_metrics, ignore_index=True)
            st.dataframe(calc_full_df.fillna(''))

            if opm_chart:
                st.plotly_chart(opm_chart, use_container_width=True)
