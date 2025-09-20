import dash
from dash import html, dcc, dash_table, Input, Output, State
import pandas as pd
from web_extract_data_from_screener_site import scrape_stock
import plotly.express as px

SECTION_NAMES = {
    'quarters': 'Qtrly. P&L',
    'profit-loss': 'Yearly P&L',
    'peers': 'Peers',
    'balance-sheet': 'Balance Sheet',
    'cash-flow': 'Cash Flow',
    'ratios': 'Ratios',
    'shareholding': 'Shareholding',
    'top-ratios': 'Top Ratios'
}

TAB_COLORS = {
    'quarters': {'backgroundColor': '#E6F2FF', 'color': '#003366'},
    'profit-loss': {'backgroundColor': '#E8F5E9', 'color': '#2E7D32'},
    'peers': {'backgroundColor': '#FFFDE7', 'color': '#FBC02D'},
    'balance-sheet': {'backgroundColor': '#F3E5F5', 'color': '#6A1B9A'},
    'cash-flow': {'backgroundColor': '#FFE0B2', 'color': '#E65100'},
    'ratios': {'backgroundColor': '#B3E5FC', 'color': '#01579B'},
    'top-ratios': {'backgroundColor': '#D1C4E9', 'color': '#4A148C'},
    'shareholding': {'backgroundColor': '#FFCDD2', 'color': '#C62828'},
    'calculated-metrics': {'backgroundColor': '#CFD8DC', 'color': '#37474F'}
}

app = dash.Dash(__name__, suppress_callback_exceptions=True)
server = app.server

app.layout = html.Div([
    html.H1("Stock Screener Data", style={'textAlign': 'center'}),
    html.Div([
        dcc.Input(id='stock-input', type='text', placeholder='Enter stock symbol (e.g., TCS)'),
        html.Button('Load Data', id='load-button', n_clicks=0)
    ], style={'textAlign': 'center', 'marginBottom': '20px'}),
    html.Div(id='status-output', style={'textAlign': 'center', 'marginBottom': '20px'}),
    dcc.Loading(
        id="loading-tabs",
        type="circle",
        children=html.Div(id='tabs-container')
    ),
    dcc.Download(id="download-excel")
])

def align_row(row, cols):
    return pd.Series([pd.to_numeric(row.get(col, None), errors='coerce') for col in cols], index=cols)

# --- Main Tabs Callback ---
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
    quarterly_metrics = []
    yearly_metrics = []
    opm_chart = None

    all_years = set()
    for tables in section_tables.values():
        for df in tables:
            if df.empty:
                continue
            df = df.copy()
            df.rename(columns={df.columns[0]: 'Parameters'}, inplace=True)
            for col in df.columns[1:]:
                if col.strip().startswith("Mar "):
                    try:
                        all_years.add(int(col.strip().split()[1]))
                    except:
                        continue
    if not all_years:
        return html.Div(), "No fiscal year columns found."
    last_8_years = sorted(all_years)[-8:]
    master_year_cols = [f"Mar {y}" for y in last_8_years]

    # placeholders
    sales_values = operating_profit_values = net_profit_values = None
    cash_eq_values = long_term_borrowings_values = short_term_borrowings_values = None
    cash_operating_values = fixed_assets_values = None
    equity_values = reserves_values = None
    top_ratios_df = None
    ratios_df = None   # <-- store Ratios
    eps_values = None

    for section_id, tables in section_tables.items():
        friendly_name = SECTION_NAMES.get(section_id, section_id)
        if not tables:
            continue

        if section_id == 'profit-loss':
            tables = tables[:1]
        if section_id == 'shareholding':
            tables = tables[:1]

        for i, df in enumerate(tables, start=1):
            if df.empty:
                continue
            df = df.copy()

            if section_id == 'peers':
                df.rename(columns={df.columns[0]: 'Company'}, inplace=True)
            else:
                df.rename(columns={df.columns[0]: 'Parameters'}, inplace=True)

            if section_id != 'quarters':
                if section_id in ['peers', 'shareholding', 'top-ratios']:
                    display_cols = df.columns
                else:
                    display_cols = ['Parameters'] + [col for col in df.columns[1:] if col in master_year_cols]
            else:
                display_cols = df.columns

            # --- RIBBON-STYLE TAB ---
            tab_style = TAB_COLORS.get(section_id, {'backgroundColor': '#f9f9f9', 'color': 'black'})
            tab_style.update({
                'height': '28px',
                'lineHeight': '28px',
                'padding': '2px 8px',
                'fontSize': '12px',
                'borderRadius': '4px',
                'margin': '2px'
            })
            selected_style = {
                'fontWeight': 'bold',
                'backgroundColor': tab_style['color'],
                'color': 'white',
                'height': '28px',
                'lineHeight': '28px',
                'padding': '2px 8px',
                'fontSize': '12px',
                'borderRadius': '4px',
                'margin': '2px'
            }

            tab_label = f"📊 {friendly_name}"
            tabs.append(dcc.Tab(
                label=tab_label,
                style={**tab_style, 'fontSize': '16px'},  # <-- increase font size here
                selected_style={**selected_style, 'fontSize': '16px'},  # <-- selected tab too
                children=[
                    dash_table.DataTable(
                        data=df[display_cols].fillna('').to_dict('records'),
                        columns=[{"name": col, "id": col} for col in display_cols],
                        page_size=len(df),
                        style_table={'overflowX': 'auto'},
                        style_header={'backgroundColor': '#003366', 'color': 'white',
                                      'fontWeight': 'bold', 'textAlign': 'center', 'height': '28px'},
                        style_cell={'textAlign': 'left', 'padding': '5px'},
                        style_data_conditional=[
                            {'if': {'column_id': display_cols[0]}, 'fontWeight': 'bold',
                             'backgroundColor': '#f2f2f2'},
                            {'if': {'row_index': 'odd'}, 'backgroundColor': '#f9f9f9'},
                        ]
                    )
                ]
            ))

            # --- Store dataframes for calculations ---
            if section_id == 'top-ratios':
                top_ratios_df = df.copy()
            if section_id == 'ratios':
                ratios_df = df.copy()

            # --- QUARTERS calculations ---
            if section_id == 'quarters':
                try:
                    quarter_cols = df.columns[1:]
                    sales_row_q = df[df['Parameters'].str.contains(r"Sales", case=False, na=False)]
                    op_row_q = df[df['Parameters'].str.contains(r"Operating Profit", case=False, na=False)]
                    interest_row_q = df[df['Parameters'].str.contains(r"Interest", case=False, na=False)]
                    if not sales_row_q.empty and not op_row_q.empty:
                        sales_q_values = align_row(sales_row_q.iloc[0], quarter_cols)
                        op_q_values = align_row(op_row_q.iloc[0], quarter_cols)
                        opm_values = (op_q_values / sales_q_values * 100).round(2).fillna(0)
                        quarterly_metrics.append(pd.DataFrame({
                            'Metric': ['Operating Profit % of Sales'],
                            **{col: [val] for col, val in zip(quarter_cols, opm_values)}
                        }))
                        if not interest_row_q.empty:
                            interest_values_q = align_row(interest_row_q.iloc[0], quarter_cols)
                            interest_cov = (op_q_values / interest_values_q).round(2).fillna(0)
                            quarterly_metrics.append(pd.DataFrame({
                                'Metric': ['Interest Coverage Ratio'],
                                **{col: [val] for col, val in zip(quarter_cols, interest_cov)}
                            }))
                        chart_df = pd.DataFrame({'Quarter': quarter_cols,
                                                 'Operating Margin (%)': opm_values.values})
                        opm_chart = dcc.Graph(
                            figure=px.bar(
                                chart_df, x='Quarter', y='Operating Margin (%)',
                                text='Operating Margin (%)',
                                color='Operating Margin (%)',
                                color_continuous_scale='Blues',
                                title='Operating Profit % of Sales (QoQ)'
                            ).update_traces(textposition='auto')
                        )
                except Exception as e:
                    print(f"Error quarterly metrics: {e}")

            # --- YEARLY METRICS ---
            if section_id == 'profit-loss':
                try:
                    sales_row = df[df['Parameters'].str.contains(r"Sales", case=False, na=False)]
                    op_row = df[df['Parameters'].str.contains(r"Operating Profit", case=False, na=False)]
                    net_row = df[df['Parameters'].str.contains(r"Net Profit", case=False, na=False)]
                    eps_row = df[df['Parameters'].str.contains(r"EPS in Rs", case=False, na=False)]
                    if not sales_row.empty:
                        sales_values = align_row(sales_row.iloc[0], master_year_cols)
                    if not op_row.empty:
                        operating_profit_values = align_row(op_row.iloc[0], master_year_cols)
                    if not net_row.empty:
                        net_profit_values = align_row(net_row.iloc[0], master_year_cols)
                    if not eps_row.empty:
                        eps_values = align_row(eps_row.iloc[0], master_year_cols)
                except Exception as e:
                    print(f"Error yearly P&L: {e}")

            # --- BALANCE SHEET ---
            if section_id == 'balance-sheet':
                try:
                    cash_eq_row = df[df['Parameters'].str.contains(r"Cash\s*Equivalents|Cash", case=False, na=False)]
                    lt_row = df[df['Parameters'].str.contains(r"Long\s*term\s*Borrowings", case=False, na=False)]
                    st_row = df[df['Parameters'].str.contains(r"Short\s*term\s*Borrowings", case=False, na=False)]
                    eq_row = df[df['Parameters'].str.contains(r"Equity", case=False, na=False)]
                    res_row = df[df['Parameters'].str.contains(r"Reserves", case=False, na=False)]
                    if not cash_eq_row.empty:
                        cash_eq_values = align_row(cash_eq_row.iloc[0], master_year_cols)
                    if not lt_row.empty:
                        long_term_borrowings_values = align_row(lt_row.iloc[0], master_year_cols)
                    if not st_row.empty:
                        short_term_borrowings_values = align_row(st_row.iloc[0], master_year_cols)
                    if not eq_row.empty:
                        equity_values = align_row(eq_row.iloc[0], master_year_cols)
                    if not res_row.empty:
                        reserves_values = align_row(res_row.iloc[0], master_year_cols)
                    if not res_row.empty:
                        reserves_values = align_row(res_row.iloc[0], master_year_cols)
                except Exception as e:
                    print(f"Error BS: {e}")

            # --- CASH FLOW ---
            if section_id == 'cash-flow':
                try:
                    cash_op_row = df[df['Parameters'].str.contains(r"Cash\s*from\s*Operating\s*Activity\s*\+", case=False, na=False)]
                    fixed_assets_row = df[df['Parameters'].str.contains(r"Fixed\s*assets\s*purchased", case=False, na=False)]
                    if not cash_op_row.empty:
                        cash_operating_values = align_row(cash_op_row.iloc[0], master_year_cols)
                    if not fixed_assets_row.empty:
                        fixed_assets_values = align_row(fixed_assets_row.iloc[0], master_year_cols)
                except Exception as e:
                    print(f"Error CashFlow: {e}")

    # ---- CALCULATED METRICS ----
    if operating_profit_values is not None and net_profit_values is not None:
        try:
            fcf_values = (cash_operating_values + fixed_assets_values).round(2) if cash_operating_values is not None and fixed_assets_values is not None else None
            tce_values = (equity_values + reserves_values + long_term_borrowings_values + short_term_borrowings_values).round(2) if all(v is not None for v in [equity_values, reserves_values, long_term_borrowings_values, short_term_borrowings_values]) else None
            roce_values = (operating_profit_values / tce_values * 100).round(2) if tce_values is not None else None
            roe_values = (net_profit_values / (equity_values + reserves_values) * 100).round(2) if equity_values is not None and reserves_values is not None else None

            ev_ebit_values = None
            if top_ratios_df is not None:
                market_cap_row = top_ratios_df[top_ratios_df['Parameters'].str.contains("Market Cap", case=False)]
                if not market_cap_row.empty:
                    market_cap = pd.to_numeric(market_cap_row.iloc[0]['Value'], errors='coerce')
                    if market_cap is not None and operating_profit_values is not None:
                        ev_ebit_values = ((market_cap + long_term_borrowings_values + short_term_borrowings_values - cash_eq_values) / operating_profit_values).round(2)

            debtor_days_values = None
            if ratios_df is not None:
                debtor_row = ratios_df[ratios_df['Parameters'].str.contains("Debtor Days", case=False, na=False)]
                if not debtor_row.empty:
                    debtor_days_values = align_row(debtor_row.iloc[0], master_year_cols)

            yearly_metrics.append(pd.DataFrame({
                'Metric': ['Sales', 'Operating Profit', 'Net Profit', 'Cash Equivalents',
                           'Long Term Borrowings', 'Short Term Borrowings',
                           'Cash from Operating Activity +', 'Fixed Assets Purchased',
                           'Free Cash Flow', 'Total Capital Employed',
                           'ROCE (%)', 'ROE (%)', 'EV/EBIT','EPS (Rs)', 'DSO/Debtor Days'],

                **{col: [
                    sales_values.get(col, None) if sales_values is not None else None,
                    operating_profit_values.get(col, None),
                    net_profit_values.get(col, None),
                    cash_eq_values.get(col, None) if cash_eq_values is not None else None,
                    long_term_borrowings_values.get(col, None) if long_term_borrowings_values is not None else None,
                    short_term_borrowings_values.get(col, None) if short_term_borrowings_values is not None else None,
                    cash_operating_values.get(col, None) if cash_operating_values is not None else None,
                    fixed_assets_values.get(col, None) if fixed_assets_values is not None else None,
                    fcf_values.get(col, None) if fcf_values is not None else None,
                    tce_values.get(col, None) if tce_values is not None else None,
                    roce_values.get(col, None) if roce_values is not None else None,
                    roe_values.get(col, None) if roe_values is not None else None,
                    ev_ebit_values.get(col, None) if ev_ebit_values is not None else None,
                    eps_values.get(col, None) if eps_values is not None else None,
                    debtor_days_values.get(col, None) if debtor_days_values is not None else None
                ] for col in master_year_cols}
            }))
        except Exception as e:
            print(f"Error computing yearly metrics: {e}")

    # ---- RENDER CALCULATED METRICS TAB ----
    tab_children = []

    if quarterly_metrics:
        calc_df_q = pd.concat(quarterly_metrics, ignore_index=True)
        # Add heading
        tab_children.append(html.H3("Quarterly Metrics", style={'textAlign': 'left', 'marginTop': '10px', 'marginBottom': '10px'}))
        tab_children.append(dash_table.DataTable(
            data=calc_df_q.fillna('').to_dict('records'),
            columns=[{"name": col, "id": col} for col in calc_df_q.columns],
            page_size=len(calc_df_q),
            style_table={'overflowX': 'auto'},
            style_header={'backgroundColor': '#004080', 'color': 'white',
                        'fontWeight': 'bold', 'textAlign': 'center'},
            style_cell={'textAlign': 'left', 'padding': '5px'},
            style_data_conditional=[
                {'if': {'column_id': 'Metric'}, 'fontWeight': 'bold', 'backgroundColor': '#e6f2ff'},
                {'if': {'row_index': 'odd'}, 'backgroundColor': '#f9f9f9'},
            ]
        ))
        if opm_chart:
            tab_children.append(html.Br())
            tab_children.append(opm_chart)


    if yearly_metrics:
        calc_df_y = pd.concat(yearly_metrics, ignore_index=True)
        tab_children.append(html.H3("Yearly Metrics (Last 8 FYs)"))
        tab_children.append(
            html.Button("Download Yearly Metrics to Excel",
                        id="download-btn",
                        n_clicks=0,
                        style={"marginBottom": "10px"})
        )
        tab_children.append(dash_table.DataTable(
            id="yearly-metrics-table",
            data=calc_df_y.fillna('').to_dict('records'),
            columns=[{"name": col, "id": col} for col in calc_df_y.columns],
            page_size=len(calc_df_y),
            style_table={'overflowX': 'auto'},
            style_header={'backgroundColor': '#004080', 'color': 'white',
                          'fontWeight': 'bold', 'textAlign': 'center','height': '28px'},
            style_cell={'textAlign': 'left', 'padding': '5px'},
            style_data_conditional=[
                {'if': {'column_id': 'Metric'}, 'fontWeight': 'bold', 'backgroundColor': '#e6f2ff'},
                {'if': {'row_index': 'odd'}, 'backgroundColor': '#f2f2f2'},
            ]
        ))

        for metric in calc_df_y['Metric']:
            row = calc_df_y[calc_df_y['Metric'] == metric]
            if not row.empty:
                values = row.iloc[0][master_year_cols].fillna(0)
                chart_df = pd.DataFrame({'Year': master_year_cols, metric: values.values})
                chart_fig = px.line(chart_df, x='Year', y=metric,
                                     markers=True, title=f"{metric} Trend (Last 8 FYs)" )
                tab_children.append(html.Br())
                tab_children.append(dcc.Graph(figure=chart_fig))

    if tab_children:
        calc_tab_style = TAB_COLORS.get('calculated-metrics', {'backgroundColor': '#CFD8DC', 'color': '#37474F'})
        calc_selected_style = {
            'fontWeight': 'bold',
            'backgroundColor': calc_tab_style['color'],
            'color': 'white'
        }

        # Add ribbon-style visual
        calc_tab_style.update({
            'fontSize': '16px',
            'padding': '4px 8px',
            'borderRadius': '6px',
            'lineHeight': '28px',  # same as other tabs
        })
        calc_selected_style.update(calc_tab_style)

        tabs.append(dcc.Tab(
            label="📈 Calculated Metrics",
            style=calc_tab_style,
            selected_style=calc_selected_style,
            children=tab_children  # no extra div
        ))



    if not tabs:
        tabs.append(dcc.Tab(label="No Data Available", children=[html.Div("No tables found.")]))

    return dcc.Tabs(tabs), status_message

# --- Download Callback ---
@app.callback(
    Output("download-excel", "data"),
    Input("download-btn", "n_clicks"),
    State("yearly-metrics-table", "data"),
    prevent_initial_call=True
)
def download_metrics(n_clicks, table_data):
    if n_clicks > 0 and table_data:
        df = pd.DataFrame(table_data)
        return dcc.send_data_frame(df.to_excel, "yearly_metrics.xlsx", index=False)
    return dash.no_update

# --- Run App ---
if __name__ == "__main__":
    import os
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 8050)),
        debug=False
    )


