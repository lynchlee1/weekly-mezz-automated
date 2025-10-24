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

from basics import parse_number, parse_date, split, get_reports_range, unpack

xlsx_columns = [
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
    '전환가액 결정방법', # 10
    '만기', #11
    'PUT', #12, empty
    '주식수', # 13
    '표면이율', # 14
    '만기이율', # 15
    '만기일', # 16
    'PUT행사일', # 17, empty
    'CALL 비중', # 18, empty
    'CALL 금리', # 19, empty
    'CALL 기한', # 20, empty
    'CALL행사일', # 21, empty
    '리픽싱조건', # 22
    '발행대상', # 23
    'Sector', # 24, empty
    '당사검토여부', # 25, empty
    '주간사', # 26
    'URL', # 27
    '리픽싱가격', #28
    '리픽싱내용', #29
    '대상주식', #30
    '옵션사항', #31
    '검산', #32
]

def table_to_xlsx(data):
    filtered_reports = data
    output_entries = []
    for report in filtered_reports:        
        all_tables = unpack(report['rcept_no'])
        first_table = None
        for table in all_tables:
            table_text = table.get_text()
            if '사채의 종류' in table_text and '권면' in table_text and '정정' not in table_text:
                first_table = table
                break

        # Refactored version : Dict-based filling
        result_dict = {}
        result_dict['종목코드'] = "A" + report['stock_code']

        result_dict['회사'] = report['corp_name']

        corp_cls = report['corp_cls']
        if corp_cls == 'Y': market = "코스피"
        elif corp_cls == 'K': market = "코스닥"
        elif corp_cls == 'N': market = "코넥스"
        else: market = "비상장"
        result_dict['상장구분'] = market

        report_nm = report['report_nm']
        if '전환' in report_nm: report_type = 'CB'
        elif '교환' in report_nm: report_type = 'EB'
        elif '신주인수권부' in report_nm: report_type = 'BW'
        else: report_type = "N/A"
        result_dict['종류'] = report_type

        result_dict['URL'] = f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={report['rcept_no']}"
        if first_table is None: continue
        for _, tr in enumerate(first_table.find_all('tr')):
            row_text = tr.get_text(' | ', strip=True)
            keyword_text = split(row_text)[0]

            if '납입일' in keyword_text:
                result_dict['납입일'] = parse_date(split(row_text)[1])

            if '사채의 종류' in keyword_text: 
                result_dict['회차'] = split(row_text)[2]
            
            if '사채의 권면' in keyword_text:
                result_dict['발행금액(억)'] = parse_number(split(row_text)[1])/10**8

            if '전환가액' in keyword_text and '원' in keyword_text:
                if '전환가액(원)' in result_dict.keys(): continue
                result_dict['전환가액(원)'] = parse_number(split(row_text)[1])
            if '교환가액' in keyword_text and '원' in keyword_text: 
                if '전환가액(원)' in result_dict.keys(): continue
                result_dict['전환가액(원)'] = parse_number(split(row_text)[1])
            if '행사가액' in keyword_text and '원' in keyword_text: 
                if '전환가액(원)' in result_dict.keys(): continue
                result_dict['전환가액(원)'] = parse_number(split(row_text)[1])

            if '전환가액 결정방법' in keyword_text: 
                result_dict['전환가액 결정방법'] = split(row_text)[1]
            elif '교환가액 결정방법' in keyword_text: 
                result_dict['전환가액 결정방법'] = split(row_text)[1]
            elif '행사가액 결정방법' in keyword_text: 
                result_dict['전환가액 결정방법'] = split(row_text)[1]
            
            if '사채의 이율' in keyword_text: 
                result_dict['표면이율'] = split(row_text)[2]
                try:
                    rate = float(split(row_text)[2].strip('%')) / 100
                    result_dict['표면이율'] = f"{round(100*rate, 1)}%"
                except:
                    None

            if '만기이자율' in keyword_text: 
                result_dict['만기이율'] = split(row_text)[1]
                try:
                    rate = float(split(row_text)[1].strip('%')) / 100
                    result_dict['만기이율'] = f"{round(100*rate, 1)}%"
                except:
                    None

            if '사채만기일' in keyword_text:
                result_dict['만기일'] = parse_date(split(row_text)[1])

            if '시가하락' in keyword_text:
                result_dict['리픽싱가격'] = parse_number(split(row_text)[2])
                if result_dict['리픽싱가격'] == -1.0: result_dict['리픽싱가격'] = "-"

            if '조정가액 근거' in keyword_text: 
                result_dict['리픽싱내용'] = ' '.join(split(row_text)[1:])
                if result_dict['리픽싱내용'] == ['-']: result_dict['리픽싱내용'] = "-"

            if '주관회사' in keyword_text:
                result_dict['주간사'] = split(row_text)[1]

            if '교환대상' in keyword_text:
                result_dict['대상주식'] = split(row_text)[2]
            elif '전환에 따라' in keyword_text:
                result_dict['대상주식'] = split(row_text)[2]

            if '옵션에 관한' in keyword_text:
                result_dict['옵션사항'] = ' '.join(split(row_text)[1:])

        try:
            date_format = "%Y-%m-%d"
            date1 = datetime.strptime(result_dict['만기일'], date_format)
            date2 = datetime.strptime(result_dict['납입일'], date_format)
            diff_days = (date1 - date2).days
            result_dict['만기'] = f"{round(diff_days/365.0, 1)}년"
        except:
            result_dict['만기'] = "-"

        try:
            result_dict['주식수'] = result_dict['발행금액(억)'] / result_dict['전환가액(원)']*10**8
        except:
            result_dict['주식수'] = "N/A"

        if '리픽싱가격' not in result_dict.keys(): result_dict['리픽싱가격'] = "-"
        if '리픽싱내용' not in result_dict.keys(): result_dict['리픽싱내용'] = "-"
        try:
            result_dict['리픽싱가격'] = f"{round(100 * result_dict['리픽싱가격'] / result_dict['전환가액(원)'],0)}%"
        except:
            result_dict['리픽싱가격'] = "-"
        
        try:
            result_dict['발행대상'], result_dict['검산'] = list_fund_participants(all_tables)
        except Exception:
            result_dict['발행대상'], result_dict['검산'] = "-", 0.0

        result_list = []
        for key in xlsx_columns:
            if key not in result_dict.keys():
                result_list.append("-")
                continue
            result_list.append(result_dict[key])
        output_entries.append(result_list)
    df = pd.DataFrame(output_entries, columns=xlsx_columns)
    
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
            '선택권', '소송', '보증', 
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
