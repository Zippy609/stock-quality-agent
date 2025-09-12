import pandas as pd
import win32com.client
import sqlite3
from write_log_to_trace_sheet import write_log_to_trace
from all_configurations import get_configurations

excel_app = win32com.client.Dispatch('Excel.Application')
source_file_path =get_configurations('main_excel_file') 
workbook = excel_app.Workbooks.Open(source_file_path)
db_path = get_configurations('db_path') 
conn = sqlite3.connect(db_path)

def empty_used_cells():
    try:
        sheet_empty = workbook.Sheets('transformed_data')
        sheet_empty.Range('J3:N7').value=""
        sheet_empty.Range('Z3:AA7').value=""
        
        sheet_empty.Range('J11:N15').value=""
        sheet_empty.Range('Z11:AA15').value=""

        sheet_empty.Range('J19:N24').value=""
        sheet_empty.Range('R19:S24').value=""

        sheet_empty.Range('J27:K31').value=""
        sheet_empty.Range('M27:M31').value=""
        sheet_empty.Range('J35:M39').value=""
        sheet_empty.Range('K42:K44').value=""
        sheet_empty.Range("N42").value=""

        # Headers
        stock_symbol = read_stock_symbol()
        sheet_empty.Range("C2").value = stock_symbol
        sheet_empty.Range("J1").value = stock_symbol + " [Annual PnL Data] updated by transform_data_from_database_to_excel.py program"
        sheet_empty.Range("J9").value = stock_symbol + " [Quarterly PnL Data] updated by transform_data_from_database_to_excel.py program"
        sheet_empty.Range("J17").value = stock_symbol + " [Balance Sheet] updated by transform_data_from_database_to_excel.py program"
        sheet_empty.Range("J25").value = stock_symbol + " [Operating Cash Flow] updated by transform_data_from_database_to_excel.py program"
        sheet_empty.Range("J33").value = stock_symbol + " [Share holding Pattern] updated by transform_data_from_database_to_excel.py program"
        sheet_empty.Range("J41").value = stock_symbol + " [Company Summary] updated by transform_data_from_database_to_excel.py program"
    except Exception as e:
        print(f"Error: {e}")

def read_stock_symbol():
    try:
        stock_worksheet = workbook.Sheets('GenRpt')
        stock_symbol = stock_worksheet.Range('D18').Value
        print("[INFO]: Transformation started for " + stock_symbol)
        return stock_symbol
        
    except Exception as e:
        print(f"Error: {e}")

write_log_to_trace("     Fetching Data from DB for Transformation for: " + read_stock_symbol())

def read_from_balance_sheet():
    try:
        balance_sheet_params = 'Equity Capital','Reserves','Long term Borrowings','Short term Borrowings','Cash Equivalents','Trade receivables'
        stock_symbol=read_stock_symbol()
        for param in balance_sheet_params:
            sql_part1= f"SELECT REPORTING_YEAR, PARAMETER_VALUE_INR_CRORE,BALANCE_SHEET_PARAMETER from BALANCE_SHEET "
            sql_where = f" WHERE STOCK_SYMBOL = '{stock_symbol}' and BALANCE_SHEET_PARAMETER = '{param}' ORDER BY REFERENCE_DATE DESC LIMIT 6"
            sql_final = sql_part1 + sql_where
            df_balance_sheet = pd.read_sql_query(sql_final,conn)
            if df_balance_sheet.empty :
                write_log_to_trace("     [ERROR] Empty Bal Sheet dataframe for " + stock_symbol + " for parameter "+param)
            write_balancesheet_to_excel(df_balance_sheet)

    except Exception as e:
        print(f"Error: {e}")

def write_balancesheet_to_excel(df_balance_sheet):
    try:

        output_worksheet = workbook.Sheets('transformed_data')
        df_balance_sheet['BAL_REPORTING_YEAR'] = "'"+df_balance_sheet['REPORTING_YEAR']
        reference_parameter = df_balance_sheet.iloc[0,2]
        df_reporting_year = pd.DataFrame(df_balance_sheet['BAL_REPORTING_YEAR'])
        df_param_value = pd.DataFrame(df_balance_sheet['PARAMETER_VALUE_INR_CRORE'])
        
        total_row_count = len(df_balance_sheet)-1
        if reference_parameter == "Equity Capital":
            excel_range_to_paste=output_worksheet.Range(f"J19:J{19+total_row_count}")
            data_list_of_reporting_year =  df_reporting_year.values.tolist()
            excel_range_to_paste.Value=data_list_of_reporting_year
        
            excel_range_to_paste=output_worksheet.Range(f"K19:K{19+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
        elif reference_parameter == "Reserves":
            excel_range_to_paste=output_worksheet.Range(f"L19:L{19+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
        elif reference_parameter == "Long term Borrowings":
            excel_range_to_paste=output_worksheet.Range(f"M19:M{19+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
        elif reference_parameter == "Short term Borrowings":
            excel_range_to_paste=output_worksheet.Range(f"N19:N{19+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
        elif reference_parameter == "Cash Equivalents":
            excel_range_to_paste=output_worksheet.Range(f"R19:R{19+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
        elif reference_parameter == "Trade receivables":
            excel_range_to_paste=output_worksheet.Range(f"S19:S{19+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
    except Exception as e:
        print(f"Error: {e}")

def read_from_pnl_qurterly_data(reporting_frequency):
    try:
        pnl_params = 'Sales','Operating Profit','Net Profit','EPS in Rs','Interest'
        stock_symbol=read_stock_symbol()
        for param in pnl_params:
            sql_part1= f"SELECT PNL_REPORTING_PERIOD, PARAMETER_VALUE_INR_CRORE,PNL_PARAMETER from PROFIT_AND_LOSS "
            sql_where = f" WHERE STOCK_SYMBOL = '{stock_symbol}' and PNL_PARAMETER = '{param}' and PNL_REPORTING_FREQUENCY = '{reporting_frequency}' ORDER BY REFERENCE_DATE DESC LIMIT 5"
            sql_final = sql_part1 + sql_where
            df_quarter_pnl = pd.read_sql_query(sql_final,conn)
            if df_quarter_pnl.empty :
                write_log_to_trace("     [ERROR] Empty Qrtrly PnL dataframe for " + stock_symbol + " for parameter "+param)
            write_quarterly_pnl_to_excel(df_quarter_pnl)

    except Exception as e:
        print(f"Error: {e}")

def write_quarterly_pnl_to_excel(df_quarter_pnl):
    try:

        output_worksheet = workbook.Sheets('transformed_data')
        df_quarter_pnl['NEW_PNL_REPORTING_PERIOD'] = "'"+df_quarter_pnl['PNL_REPORTING_PERIOD']
        reference_parameter = df_quarter_pnl.iloc[0,2]
        df_reporting_year = pd.DataFrame(df_quarter_pnl['NEW_PNL_REPORTING_PERIOD'])
        df_param_value = pd.DataFrame(df_quarter_pnl['PARAMETER_VALUE_INR_CRORE'])
        
        total_row_count = len(df_quarter_pnl)-1
        if reference_parameter == "Sales":
            excel_range_to_paste=output_worksheet.Range(f"J11:J{11+total_row_count}")
            data_list_of_reporting_year =  df_reporting_year.values.tolist()
            excel_range_to_paste.Value=data_list_of_reporting_year
            
            excel_range_to_paste=output_worksheet.Range(f"K11:K{11+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
        elif reference_parameter == "Operating Profit":
                excel_range_to_paste=output_worksheet.Range(f"M11:M{11+total_row_count}")
                excel_range_to_paste.Value=df_param_value.values.tolist()
        elif reference_parameter == "Net Profit":
                excel_range_to_paste=output_worksheet.Range(f"N11:N{11+total_row_count}")
                excel_range_to_paste.Value=df_param_value.values.tolist()
        elif reference_parameter == "EPS in Rs":
                excel_range_to_paste=output_worksheet.Range(f"Z11:Z{11+total_row_count}")
                excel_range_to_paste.Value=df_param_value.values.tolist()
        elif reference_parameter == "Interest":
                excel_range_to_paste=output_worksheet.Range(f"AA11:AA{11+total_row_count}")
                excel_range_to_paste.Value=df_param_value.values.tolist()
        #write_log_to_trace("     Transformed Qrtly PnL for " + reference_parameter)
    except Exception as e:
        print(f"Error: {e}")

def read_from_pnl_yearly_data(reporting_frequency):
    try:
        pnl_params = 'Sales','Operating Profit','Net Profit','EPS in Rs','Interest'
        stock_symbol=read_stock_symbol()
        for param in pnl_params:
            sql_part1= f"SELECT PNL_REPORTING_PERIOD, PARAMETER_VALUE_INR_CRORE,PNL_PARAMETER from PROFIT_AND_LOSS "
            sql_where = f" WHERE STOCK_SYMBOL = '{stock_symbol}' and PNL_PARAMETER = '{param}' and PNL_REPORTING_FREQUENCY = '{reporting_frequency}' ORDER BY REFERENCE_DATE DESC LIMIT 5"
            sql_final = sql_part1 + sql_where
            df_yearly_pnl = pd.read_sql_query(sql_final,conn)
            if df_yearly_pnl.empty :
                write_log_to_trace("     [ERROR] Empty Yearly PnL dataframe for " + stock_symbol + " for parameter "+param)

            write_yearly_pnl_to_excel(df_yearly_pnl)

    except Exception as e:
        print(f"Error: {e}")

def write_yearly_pnl_to_excel(df_yearly_pnl):
    try:

        output_worksheet = workbook.Sheets('transformed_data')
        df_yearly_pnl['NEW_PNL_REPORTING_PERIOD'] = "'"+df_yearly_pnl['PNL_REPORTING_PERIOD']
        reference_parameter = df_yearly_pnl.iloc[0,2]
        df_reporting_year = pd.DataFrame(df_yearly_pnl['NEW_PNL_REPORTING_PERIOD'])
        df_param_value = pd.DataFrame(df_yearly_pnl['PARAMETER_VALUE_INR_CRORE'])
        
        total_row_count = len(df_yearly_pnl)-1

        if reference_parameter == "Sales":
            excel_range_to_paste=output_worksheet.Range(f"J3:J{3+total_row_count}")
            data_list_of_reporting_year =  df_reporting_year.values.tolist()
            excel_range_to_paste.Value=data_list_of_reporting_year
        
            excel_range_to_paste=output_worksheet.Range(f"K3:K{3+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
        elif reference_parameter == "Operating Profit":
            excel_range_to_paste=output_worksheet.Range(f"M3:M{3+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
        elif reference_parameter == "Net Profit":
            excel_range_to_paste=output_worksheet.Range(f"N3:N{3+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
        elif reference_parameter == "EPS in Rs":
            excel_range_to_paste=output_worksheet.Range(f"Z3:Z{3+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
        elif reference_parameter == "Interest":
            excel_range_to_paste=output_worksheet.Range(f"AA3:AA{3+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
    except Exception as e:
        print(f"Error: {e}")

def read_from_cash_flow_data():
    try:
        cash_flow_params = 'Cash from Operating Activity','Fixed assets purchased'
        stock_symbol=read_stock_symbol()
        for param in cash_flow_params:
            sql_part1= f"SELECT CASH_FLOW_PERIOD, PARAMETER_VALUE_INR_CRORE,CASH_FLOW_PARAMETER from CASH_FLOW_STATEMENT "
            sql_where = f" WHERE STOCK_SYMBOL = '{stock_symbol}' and CASH_FLOW_PARAMETER = '{param}' ORDER BY REFERENCE_DATE DESC LIMIT 5"
            sql_final = sql_part1 + sql_where
            df_cash_flow = pd.read_sql_query(sql_final,conn)
            if df_cash_flow.empty :
                write_log_to_trace("     [ERROR] Empty Cash flow dataframe for " + stock_symbol + " for parameter "+cash_flow_params)

            write_cash_flow_to_excel(df_cash_flow)

    except Exception as e:
        print(f"Error: {e}")

def write_cash_flow_to_excel(df_cash_flow):
    try:
        output_worksheet = workbook.Sheets('transformed_data')
        df_cash_flow['NEW_CASH_FLOW_PERIOD'] = "'"+df_cash_flow['CASH_FLOW_PERIOD']
        reference_parameter = df_cash_flow.iloc[0,2]
        df_reporting_year = pd.DataFrame(df_cash_flow['NEW_CASH_FLOW_PERIOD'])
        df_param_value = pd.DataFrame(df_cash_flow['PARAMETER_VALUE_INR_CRORE'])
        
        total_row_count = len(df_cash_flow)-1
        if reference_parameter == "Cash from Operating Activity":
            excel_range_to_paste=output_worksheet.Range(f"J27:J{27+total_row_count}")
            data_list_of_reporting_year =  df_reporting_year.values.tolist()
            excel_range_to_paste.Value=data_list_of_reporting_year
        
            excel_range_to_paste=output_worksheet.Range(f"K27:K{27+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
        if reference_parameter == "Fixed assets purchased":
            excel_range_to_paste=output_worksheet.Range(f"M27:M{27+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()

    except Exception as e:
        print(f"Error: {e}")

def read_from_shareholding():
    try:
        shareholding_params = 'Promoters','FIIs','DIIs'
        stock_symbol=read_stock_symbol()
        for param in shareholding_params:
            sql_part1= f"SELECT SHARE_HOLDING_PERIOD, SHARE_HOLDING_PERCENTAGE,SHARE_HOLDERS from SHARE_HOLDING_PATTERN "
            sql_where = f" WHERE STOCK_SYMBOL = '{stock_symbol}' and SHARE_HOLDERS = '{param}' ORDER BY REFERENCE_DATE DESC LIMIT 5"
            sql_final = sql_part1 + sql_where
            df_shareholding = pd.read_sql_query(sql_final,conn)
            if df_shareholding.empty :
                write_log_to_trace("     [ERROR] Empty Shareholding dataframe for " + stock_symbol + " for parameter "+param)

            write_shareholding_to_excel(df_shareholding)

    except Exception as e:
        print(f"Error: {e}")

def write_shareholding_to_excel(df_shareholding):
    try:

        output_worksheet = workbook.Sheets('transformed_data')
        df_shareholding['NEW_SHARE_HOLDING_PERIOD'] = "'"+df_shareholding['SHARE_HOLDING_PERIOD']
        reference_parameter = df_shareholding.iloc[0,2]
        df_reporting_year = pd.DataFrame(df_shareholding['SHARE_HOLDING_PERIOD'])
        df_param_value = pd.DataFrame(df_shareholding['SHARE_HOLDING_PERCENTAGE'])/100
        
        total_row_count = len(df_shareholding)-1
        if reference_parameter == "Promoters":
            excel_range_to_paste=output_worksheet.Range(f"J35:J{35+total_row_count}")
            data_list_of_reporting_year =  df_reporting_year.values.tolist()
            excel_range_to_paste.Value=data_list_of_reporting_year
        
            excel_range_to_paste=output_worksheet.Range(f"K35:K{35+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
        elif reference_parameter == "FIIs":
            excel_range_to_paste=output_worksheet.Range(f"L35:L{35+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
        elif reference_parameter == "DIIs":
            excel_range_to_paste=output_worksheet.Range(f"M35:M{35+total_row_count}")
            excel_range_to_paste.Value=df_param_value.values.tolist()
    except Exception as e:
        print(f"Error: {e}")

def read_from_company_summary():
    try:
        stock_symbol=read_stock_symbol()
        sql_part1= f"SELECT MARKET_CAP,CURRENT_PRICE,MEDIAN_PE_FROM_PEER_GROUP FROM COMPANY_SUMMARY "
        sql_where = f" WHERE STOCK_SYMBOL = '{stock_symbol}' "
        sql_final = sql_part1 + sql_where
        df_com_summary = pd.read_sql_query(sql_final,conn)
        if df_com_summary.empty :
            write_log_to_trace("     [ERROR] Empty Company Summary dataframe for " + stock_symbol)

        output_worksheet = workbook.Sheets('transformed_data')
        market_cap = df_com_summary.iloc[0,0]
        current_price=df_com_summary.iloc[0,1]
        median_pe_from_perr_group = df_com_summary.iloc[0,2]
        excel_range_to_paste=output_worksheet.Range(f"K42:K42")
        excel_range_to_paste.Value=market_cap.replace('₹ ','').replace(' Cr.','')
        excel_range_to_paste=output_worksheet.Range(f"K43:K43")
        excel_range_to_paste.Value=current_price.replace('₹ ','')
        excel_range_to_paste=output_worksheet.Range(f"N42:N42")
        excel_range_to_paste.Value=median_pe_from_perr_group

    except Exception as e:
        print(f"Error: {e}")



if __name__ == "__main__":
    empty_used_cells()
    read_from_balance_sheet()
    read_from_pnl_qurterly_data('Quarterly')
    read_from_pnl_yearly_data('Yearly')
    read_from_cash_flow_data()
    read_from_shareholding()
    read_from_company_summary()
    conn.close()

