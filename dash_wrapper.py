import dash
from dash import html, dcc, dash_table, Input, Output, State
import pandas as pd
from extract_data_from_screener_site import scrape_stock  # scraper program
import plotly.express as px
import os

SECTION_NAMES = {
    'quarters': 'Quarterly P&L',
    'profit-loss': 'Yearly P&L',
    'peers': 'Peers',
    'balance-sheet': 'Balance Sheet',
    'cash-flow': 'Cash Flow',
    'ratios': 'Ratios',
    'shareholding': 'Shareholding'
}

app = dash.Dash(__name__)
server = app.server  # Required for Render

app.layout = html.Div([
    html.H1("Stock Screener Data", style={'textAlign': 'center'}),
    html.Div([
        dcc.Input(id='stock-input', type='text', placeholder='Enter stock symbol (e.g., TCS)'),
        html.Button('Load Data', id='load-button', n_clicks=0)
    ], style={'textAlign': 'center', 'marginBottom': '20px'}),
    html.Div(id='status-output', style={'textAlign': 'center', 'marginBottom': '20px'}),
    html.Div(id='tabs-container')
])

@app.callback(
    Output('tabs-container', 'children'),
    Output('status-output', 'children'),
    Input('load-button', 'n_clicks'),
    State('stock-input', 'value')
)
def update_tabs(n_clicks, stock_symbol):
    if not stock_symbol:
        return html.Div(), "Please enter a stock symbol."

    try:
        status_message, section_tables = scrape_stock(stock_symbol.upper())
    except Exception as e:
        return html.Div(), f"Error scraping stock: {e}"

    tabs = []
    calculated_metrics = []
    opm_chart = None

    for section_id, tables in section_tables.items():
        friendly_name = SECTION_NAMES.get(section_id, section_id)
        if not tables:
            continue
        for i, df in enumerate(tables, start=1):
            if df.empty:
                continue

            df = df.copy()

            tab_label = f"📊 {friendly_name} Table {i}"
            tabs.append(
                dcc.Tab(
                    label=tab_label,
                    children=[
                        dash_table.DataTable(
                            data=df.fillna('').to_dict('records'),
                            columns=[{"name": col, "id": col} for col in df.columns],
                            page_size=10,
                            style_table={'overflowX': 'auto'},
                            style_header={'backgroundColor': '#003366', 'color': 'white', 'fontWeight': 'bold', 'textAlign': 'center'},
                            style_cell={'textAlign': 'left', 'padding': '5px'},
                            style_data_conditional=[
                                {'if': {'column_id': 'Parameters'}, 'fontWeight': 'bold', 'backgroundColor': '#f2f2f2'},
                                {'if': {'row_index': 'odd'}, 'backgroundColor': '#f9f9f9'},
                            ]
                        )
                    ]
                )
            )

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
                        calculated_metrics.append(pd.DataFrame({
                            'Metric': ['Operating Profit % of Sales'],
                            **{col: [val] for col, val in zip(quarter_cols, opm_values)}
                        }))

                        chart_df = pd.DataFrame({
                            'Quarter': quarter_cols,
                            'Operating Margin (%)': opm_values.values
                        })
                        opm_chart = dcc.Graph(
                            figure=px.bar(
                                chart_df,
                                x='Quarter',
                                y='Operating Margin (%)',
                                text='Operating Margin (%)',
                                color='Operating Margin (%)',
                                color_continuous_scale='Blues',
                                title='Operating Profit % of Sales (QoQ)'
                            ).update_traces(textposition='auto')
                        )
                except Exception as e:
                    print(f"Error calculating QoQ Operating Margin: {e}")

    # Calculated Metrics Tab
    if calculated_metrics:
        calc_df = pd.concat(calculated_metrics, ignore_index=True)
        tab_children = [
            dash_table.DataTable(
                data=calc_df.fillna('').to_dict('records'),
                columns=[{"name": col, "id": col} for col in calc_df.columns],
                page_size=10,
                style_table={'overflowX': 'auto'},
                style_header={'backgroundColor': '#004080', 'color': 'white', 'fontWeight': 'bold', 'textAlign': 'center'},
                style_cell={'textAlign': 'left', 'padding': '5px'},
                style_data_conditional=[
                    {'if': {'column_id': 'Metric'}, 'fontWeight': 'bold', 'backgroundColor': '#e6f2ff'},
                    {'if': {'row_index': 'odd'}, 'backgroundColor': '#f2f2f2'},
                ]
            )
        ]
        if opm_chart:
            tab_children.append(html.Br())
            tab_children.append(opm_chart)

        tabs.append(dcc.Tab(
            label="📈 Calculated Metrics",
            children=tab_children
        ))

    if not tabs:
        tabs.append(dcc.Tab(label="No Data Available", children=[html.Div("No tables found.")]))

    return dcc.Tabs(tabs), status_message

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8050))
    app.run_server(host="0.0.0.0", port=port, debug=True)
