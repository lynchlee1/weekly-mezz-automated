# We'll load API_KEY dynamically to allow for runtime updates
from datetime import datetime, timezone, timedelta
import requests
from bs4 import BeautifulSoup
import zipfile
from io import BytesIO
import pandas as pd
import sys
import os
import json
from sub import list_fund_participants

# Suppress BeautifulSoup XML parsing warnings
from bs4 import XMLParsedAsHTMLWarning
import warnings
warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

from basics import *

def table_to_xlsx(data):
    filtered_reports = data
    output_entries = []
    for report in filtered_reports:
        stock_code = report['stock_code']
        report_nm = report['report_nm']
        corp_code = report['corp_code']
        corp_name = report['corp_name']
        rcept_no = report['rcept_no']
        corp_cls = report['corp_cls']

        if '전환' in report_nm: report_type = 'CB'
        elif '교환' in report_nm: report_type = 'EB'
        elif '신주인수권부' in report_nm: report_type = 'BW'
        else: continue
        
        all_tables = unpack(rcept_no)
        first_table = None
        for table in all_tables:
            table_text = table.get_text()
            if '사채의 종류' in table_text and '권면' in table_text and '정정' not in table_text:
                first_table = table
                break
    
        if first_table is not None:
            # Initialize the result list
            result_list = [""] * 32  # 0-31 indices
            result_list[0] = "A" + stock_code
            result_list[2] = corp_name

            if corp_cls == 'Y': market = "코스피"
            elif corp_cls == 'K': market = "코스닥"
            elif corp_cls == 'N': market = "코넥스"
            elif corp_cls == 'E': market = "기타"
            else: market = "미분류"
            result_list[3] = market

            if '전환' in report_nm: report_type = 'CB'
            elif '교환' in report_nm: report_type = 'EB'
            else: report_type = 'BW'
            result_list[4] = report_type 

            result_list[27] = f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcept_no}"

            # Collect all rows and categorize them
            for row_idx, tr in enumerate(first_table.find_all('tr')):
                row_text = tr.get_text(' | ', strip=True)
                keyword_text = split(row_text)[0]

                if '납입일' in keyword_text: result_list[1] = parse_date(split(row_text)[1])
                if '사채의 종류' in keyword_text: result_list[5] = split(row_text)[2]
                if '사채의 권면' in keyword_text: result_list[8] = parse_number(split(row_text)[1])/10**8

                if '전환가액' in keyword_text and '원' in keyword_text: 
                    if result_list[9] != "": continue
                    result_list[9] = parse_number(split(row_text)[1])
                elif '교환가액' in keyword_text and '원' in keyword_text: 
                    if result_list[9] != "": continue
                    result_list[9] = parse_number(split(row_text)[1])
                elif '행사가액' in keyword_text and '원' in keyword_text: 
                    if result_list[9] != "": continue
                    result_list[9] = parse_number(split(row_text)[1])
                
                if '전환가액 결정방법' in keyword_text: result_list[10] = split(row_text)[1]
                elif '교환가액 결정방법' in keyword_text: result_list[10] = split(row_text)[1]
                elif '행사가액 결정방법' in keyword_text: result_list[10] = split(row_text)[1]

                if '사채의 이율' in keyword_text: result_list[14] = split(row_text)[2]
                if '만기이자율' in keyword_text: result_list[15] = split(row_text)[1]
                if '사채만기일' in keyword_text: result_list[16] = parse_date(split(row_text)[1])
                
                if '시가하락' in keyword_text: 
                    result_list[28] = parse_number(split(row_text)[2])
                    if result_list[28] == -1.0: result_list[28] = "-"
                if '조정가액 근거' in keyword_text: result_list[29] = split(row_text)[1:]

                if '주관회사' in keyword_text: result_list[26] = split(row_text)[1]
                if '교환대상' in keyword_text: result_list[30] = split(row_text)[2]
                elif '전환에 따라' in keyword_text: result_list[30] = split(row_text)[2]

                if '옵션에 관한' in keyword_text: result_list[31] = ' | '.join(split(row_text)[1:])

            try:
                date_format = "%Y-%m-%d"
                date1 = datetime.strptime(result_list[16], date_format)
                date2 = datetime.strptime(result_list[1], date_format)
                diff_days = (date1 - date2).days
                result_list[11] = f"{round(diff_days/365.0, 1)}년"
            except:
                result_list[11] = "-"
            try:
                result_list[13] = result_list[8] / result_list[9]*10**8
            except:
                result_list[13] = "N/A"

            if result_list[28] == "": result_list[28] = "-"
            if result_list[29] == "": result_list[29] = "-"
            try:
                if result_list[28] == -1.0: result_list[22] = "-"
                else: result_list[22] = f"{round(100 * result_list[28] / result_list[9],0)}%"
            except:
                result_list[22] = "-"
            
            try:
                result_list[23] = list_fund_participants(all_tables)
            except Exception:
                result_list[23] = "-"
            
            output_entries.append(result_list)

    # Define column names for output if desired
    columns = [
        '종목코드', # 0
        '납입일', # 1
        '회사', # 2
        '상장구분', # 3
        '종류', # 4
        '회차', # 5
        '벤처여부', # 6, empty
        '시가총액', # 7, empty
        '발행금액(억)', # 8
        '전환가액(원)', # 9
        '전환/교환가액 결정방법', # 10
        '만기', #11
        'PUT', #12, empty
        '주식수', # 13
        '표면이율', # 14
        '만기이율', # 15
        '만기', # 16
        'PUT행사일', # 17, empty
        'CALL 비중', # 18, empty
        'CALL 금리', # 19, empty
        'CALL 기한', # 20, empty
        'CALL행사일', # 21, empty
        '리픽싱조건', # 22
        '발행대상', # 23, empty
        'Sector', # 24, empty
        '당사검토여부', # 25, empty
        '주간사', # 26
        'URL', # 27
        '리픽싱가격', #28
        '리픽싱내용', #29
        '대상주식', #30
        '옵션사항', #31
    ]

    df = pd.DataFrame(output_entries, columns=columns)
    
    # Get the directory where the executable is located
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        exe_dir = os.path.dirname(sys.executable)
    else:
        # Running as script
        exe_dir = os.path.dirname(os.path.abspath(__file__))
    
    output_path = os.path.join(exe_dir, 'output.xlsx')
    df.to_excel(output_path, index=False)
    print(f"Saved results to {output_path}")

if __name__ == "__main__":
    if len(sys.argv) == 3:
        from_date = sys.argv[1]
        to_date = sys.argv[2]
        print(f"Using dates from command line: {from_date} to {to_date}")
        
        reports = get_reports_range(from_date, to_date)

        filter_words = [
            '감자', '증자', '합병', '분할', '해산', '증여',
            '자기', '자본', '자산', '담보', 
            '양수도', '양수', '양도', '처분', 
            '선택권', '소조',  '보증', 
        ]

        filtered_reports = []
        for report in reports:
            if any(word in report['report_nm'] for word in filter_words): continue
            contain_keys = ['stock_code', 'report_nm', 'corp_code', 'corp_name', 'rcept_no', 'corp_cls']
            filtered_report = {key: report[key] for key in contain_keys if key in report}
            filtered_reports.append(filtered_report)
        table_to_xlsx(filtered_reports)
    
    else:
        try:
            import tkinter as tk
            from tkinter import ttk
            from gui import DateRangeGUI
            
            root = tk.Tk()
            style = ttk.Style()
            style.theme_use('clam')
            
            app = DateRangeGUI(root)
            root.mainloop()
            
        except ImportError as e:
            print(f"Error launching GUI: {e}")
            sys.exit(1)
