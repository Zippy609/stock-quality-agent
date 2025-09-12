# file: dash_app.py

import dash
from dash import html, dcc, Input, Output, dash_table
import dash_bootstrap_components as dbc
import pandas as pd
import requests


# Create Dash app
#app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])


# Layout with input and tabs
app.layout = html.Div([
    html.H1("Indian Stock Financials Dashboard"),
    
    html.Label("Enter Stock Symbol:"),
    dcc.Input(id='symbol-input', value='TCS', type='text'),
    html.Button("Fetch Data", id='fetch-btn', n_clicks=0),
    
    dcc.Tabs(id='tabs', value='balance-sheet', children=[
        dcc.Tab(label='Balance Sheet', value='balance-sheet'),
        dcc.Tab(label='Profit & Loss', value='pl'),
        dcc.Tab(label='Cash Flow', value='cash-flow'),
        dcc.Tab(label='Quarters', value='quarters'),
        dcc.Tab(label='Peers', value='peers')
    ]),
    
    html.Div(id='tab-content')
])

# Callback to update tab content
@app.callback(
    Output('tab-content', 'children'),
    Input('fetch-btn', 'n_clicks'),
    Input('symbol-input', 'value'),
    Input('tabs', 'value')
)
def update_tab(n_clicks, symbol, tab):
    if not symbol:
        return "Enter a stock symbol"

    url = f"http://127.0.0.1:5000/get_financials?symbol={symbol.upper()}"
    try:
        data = requests.get(url).json()['financials']
    except:
        return "Error fetching data from API"

    section_map = {
        'balance-sheet': 'balance_sheet',
        'pl': 'profit_loss',
        'cash-flow': 'cash_flow',
        'quarters': 'quarters',
        'peers': 'peers'
    }

    section = data.get(section_map[tab], None)
    if not section:
        return "No data available"

    # Convert to DataFrame for Dash table
    df = pd.DataFrame(section['rows'], columns=section['headers'])
    table = dash_table.DataTable(
        data=df.to_dict('records'),
        columns=[{"name": i, "id": i} for i in df.columns],
        style_table={'overflowX': 'auto'},
        style_cell={'textAlign': 'center'},
        style_header={'fontWeight': 'bold'}
    )
    return table

if __name__ == '__main__':
    app.run(debug=True)

