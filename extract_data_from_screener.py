from flask import Flask, request, jsonify
import requests
from bs4 import BeautifulSoup

app = Flask(__name__)

BASE_URL = "https://www.screener.in/company/{symbol}/consolidated/"

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

def fetch_page(symbol, section):
    url = BASE_URL.format(symbol=symbol) + section
    response = requests.get(url, headers=HEADERS)
    if response.status_code == 200:
        return response.text
    return None

def parse_table(html, section_id):
    soup = BeautifulSoup(html, 'html.parser')
    section = soup.find("section", id=section_id)
    if not section:
        return "Section not found"
    
    table = section.find("table")
    if not table:
        return "Table not found"

    headers = [th.get_text(strip=True) for th in table.find("thead").find_all("th")]
    rows_data = []
    for row in table.find("tbody").find_all("tr"):
        cols = [td.get_text(strip=True) for td in row.find_all("td")]
        if cols:
            rows_data.append(cols)
    return {
        "headers": headers,
        "rows": rows_data
    }

@app.route('/get_financials', methods=['GET'])
def get_financials():
    symbol = request.args.get('symbol', '').upper()
    if not symbol:
        return jsonify({"error": "Please provide a stock symbol"}), 400
    
    sections = {
        "quarters": "quarters",
        "peers": "peers",
        "profit_loss": "profit-loss",
        "balance_sheet": "balance-sheet",
        "cash_flow": "cash-flow"
    }

    result = {}
    for key, section_id in sections.items():
        html = fetch_page(symbol, f"#{section_id}")
        if html:
            data = parse_table(html, section_id)
            result[key] = data
        else:
            result[key] = "Failed to fetch data"

    return jsonify({
        "symbol": symbol,
        "financials": result
    })

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
