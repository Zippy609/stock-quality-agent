import os
from io import StringIO
from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time


def extract_tables_in_section(html_content, section_id):
    soup = BeautifulSoup(html_content, 'html.parser')
    section = soup.find(id=section_id)
    if section:
        tables = section.find_all('table')
        return tables
    return []


def tables_to_dataframes(tables):
    dfs = []
    for table in tables:
        html_string = str(table)
        html_io = StringIO(html_string)
        df = pd.read_html(html_io)[0]
        dfs.append(df)
    return dfs


def click_button(driver, button_xpath):
    try:
        button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, button_xpath))
        )
        button.click()
        return True
    except Exception as e:
        print(f"Error clicking button '{button_xpath}': {e}")
        return False


def clean_value(value):
    """Remove currency, commas, percent, and units like Cr."""
    value = (
        value.replace("₹", "")
        .replace("Cr.", "")
        .replace(",", "")
        .replace("%", "")
        .strip()
    )
    return value


def extract_top_ratios(html_content):
    """Extract key ratios under <ul id="top-ratios"> and return as DataFrame"""
    soup = BeautifulSoup(html_content, 'html.parser')
    ul = soup.find("ul", id="top-ratios")
    ratios = []

    if ul:
        for li in ul.find_all("li"):
            text = li.get_text(" ", strip=True)

            # Market Cap
            if "Market Cap" in text:
                key = "Market Cap"
                value = text.split("Market Cap", 1)[1].strip()
                ratios.append({"Metric": key, "Value": clean_value(value)})
                continue

            # Current Price
            if "Current Price" in text:
                key = "Current Price"
                value = text.split("Current Price", 1)[1].strip()
                ratios.append({"Metric": key, "Value": clean_value(value)})
                continue

            # High / Low
            if "High / Low" in text:
                key = "High / Low"
                value = text.split("High / Low", 1)[1].strip()
                if "/" in value:
                    high, low = value.split("/", 1)
                    ratios.append({"Metric": "High", "Value": clean_value(high)})
                    ratios.append({"Metric": "Low", "Value": clean_value(low)})
                else:
                    ratios.append({"Metric": key, "Value": clean_value(value)})
                continue

            # Everything else (split by last space → key, value)
            parts = text.rsplit(" ", 1)
            if len(parts) == 2:
                key, value = parts
                ratios.append({"Metric": key.strip(), "Value": clean_value(value)})
            else:
                ratios.append({"Metric": text.strip(), "Value": ""})

    if ratios:
        return pd.DataFrame(ratios)
    return pd.DataFrame()


def scrape_stock(stock_name="TCS"):
    """
    Returns:
    - status message string
    - section_tables: dict of section_id -> list of DataFrames
    """
    results = []
    section_tables = {}

    if stock_name:
        url = f"https://www.screener.in/company/{stock_name}/consolidated/"
        results.append(f'[INFO]: URL: {url}')

        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.binary_location = "/usr/bin/chromium-browser"

        # Use system chromedriver (installed from apt.txt)
        service = Service("/usr/bin/chromedriver")
        driver = webdriver.Chrome(service=service, options=options)
        driver.get(url)

        # Click expand buttons if available
        for btn_text in ["Borrowings", "Other Assets", "Cash from Investing Activity"]:
            button_xpath = f'//button[contains(text(), "{btn_text}")]'
            click_button(driver, button_xpath)

        time.sleep(5)
        html_content = driver.page_source

        # Extract standard sections
        section_ids = [
            'quarters',
            'profit-loss',
            'peers',
            'balance-sheet',
            'cash-flow',
            'ratios',
            'shareholding'
        ]
        results.append(f"Section IDs for scraping: {section_ids}")

        for each_section_id in section_ids:
            tables = extract_tables_in_section(html_content, each_section_id)
            if tables:
                dfs = tables_to_dataframes(tables)
                for df in dfs:
                    if 'Unnamed: 0' in df.columns:
                        df.rename(columns={'Unnamed: 0': 'Parameters'}, inplace=True)
                section_tables[each_section_id] = dfs
            else:
                section_tables[each_section_id] = []

        # ✅ Extract Top Ratios separately
        top_ratios_df = extract_top_ratios(html_content)
        if not top_ratios_df.empty:
            section_tables["top-ratios"] = [top_ratios_df]
        else:
            section_tables["top-ratios"] = []

        driver.quit()
    return results[0], section_tables


if __name__ == "__main__":
    status, data = scrape_stock("TCS")
    print(status)
    for k, v in data.items():
        print(f"Section: {k}, Tables found: {len(v)}")
        if v:
            print(v[0].head())
