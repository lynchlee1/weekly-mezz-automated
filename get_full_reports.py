import json
import os
import zipfile
from io import BytesIO

import requests
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
import warnings
warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

from config import API_KEY
from basics import parse_number, parse_date, split
from sub import list_fund_participants
from datetime import datetime

GROUPED_FILE = "filtered_B001_list_grouped.json"
DOCUMENT_URL = "https://opendart.fss.or.kr/api/document.xml"
OUTPUT_DIR = "dart_documents"
BATCH_SIZE = 50

def unpack(rcept_no: str) -> list:
    '''
    input: rcept_no
    output: list of tables in the report
    '''
    source_download_url = "https://opendart.fss.or.kr/api/document.xml"
    url = f"{source_download_url}?crtfc_key={API_KEY}&rcept_no={rcept_no}"
    response = requests.get(url)
    
    all_tables = []
    try:
        with zipfile.ZipFile(BytesIO(response.content)) as zf:
            file_list = zf.namelist()
            for name in file_list:
                with zf.open(name) as f:
                    data = f.read()
                    try: text = data.decode('utf-8')
                    except UnicodeDecodeError:
                        # print(f"Error decoding with utf-8: {name}")
                        try: text = data.decode('cp949')
                        except UnicodeDecodeError: 
                            # print(f"- Error decoding with cp949: {name}")
                            print("Decoding Error")
                            text = None

                    if text is not None:
                        soup = BeautifulSoup(text, 'html.parser')
                        tables = soup.find_all('table')
                        all_tables.extend(tables)
    except zipfile.BadZipFile: pass
    return all_tables

def extract_table_data(report: dict, tables: list) -> dict:
    target_table = None
    for table in tables:
        table_text = table.get_text()
        if "사채의 종류" in table_text and "권면" in table_text and "정정" not in table_text:
            target_table = table
            break
    if target_table is None: return {}

    result_dict = {}

    report_nm = report.get('report_nm', '')
    if '전환사채' in report_nm: report_type = 'CB'
    elif '교환사채' in report_nm: report_type = 'EB'
    elif '신주인수권부사채' in report_nm: report_type = 'BW'
    else: report_type = "N/A"
    result_dict['종류'] = report_type

    for _, tr in enumerate(target_table.find_all('tr')):
        row_text = tr.get_text(' | ', strip=True)
        row_parts = split(row_text)
        if len(row_parts) == 0: continue
        keyword_text = row_parts[0]

        if '납입일' in keyword_text:
            if len(row_parts) > 1:
                result_dict['납입일'] = parse_date(row_parts[1])

        if '사채의 종류' in keyword_text: 
            if len(row_parts) > 2:
                result_dict['회차'] = row_parts[2]
        
        if '사채의 권면' in keyword_text:
            if len(row_parts) > 1:
                result_dict['발행금액(억)'] = parse_number(row_parts[1])/10**8

        if '전환가액' in keyword_text and '원' in keyword_text:
            if '전환가액(원)' in result_dict.keys(): continue
            if len(row_parts) > 1:
                result_dict['전환가액(원)'] = parse_number(row_parts[1])
        if '교환가액' in keyword_text and '원' in keyword_text: 
            if '전환가액(원)' in result_dict.keys(): continue
            if len(row_parts) > 1:
                result_dict['전환가액(원)'] = parse_number(row_parts[1])
        if '행사가액' in keyword_text and '원' in keyword_text: 
            if '전환가액(원)' in result_dict.keys(): continue
            if len(row_parts) > 1:
                result_dict['전환가액(원)'] = parse_number(row_parts[1])

        if '전환가액 결정방법' in keyword_text: 
            if len(row_parts) > 1: result_dict['전환가액 결정방법'] = row_parts[1:]
        elif '교환가액 결정방법' in keyword_text: 
            if len(row_parts) > 1: result_dict['전환가액 결정방법'] = row_parts[1:]
        elif '행사가액 결정방법' in keyword_text: 
            if len(row_parts) > 1: result_dict['전환가액 결정방법'] = row_parts[1:]
        
        if '사채의 이율' in keyword_text:
            if len(row_parts) > 2:
                result_dict['표면이율'] = row_parts[2]
                try:
                    rate = float(row_parts[2].strip('%')) / 100
                    result_dict['표면이율'] = f"{round(100*rate, 1)}%"
                except:
                    None

        if '만기이자율' in keyword_text: 
            if len(row_parts) > 1:
                result_dict['만기이율'] = row_parts[1]
                try:
                    rate = float(row_parts[1].strip('%')) / 100
                    result_dict['만기이율'] = f"{round(100*rate, 1)}%"
                except:
                    None

        if '사채만기일' in keyword_text:
            if len(row_parts) > 1:
                result_dict['만기일'] = parse_date(row_parts[1])

        if '시가하락' in keyword_text:
            if len(row_parts) > 2:
                result_dict['리픽싱가격'] = parse_number(row_parts[2])
                if result_dict['리픽싱가격'] == -1.0: result_dict['리픽싱가격'] = "-"

        if '조정가액 근거' in keyword_text: 
            if len(row_parts) > 1:
                result_dict['리픽싱내용'] = ' '.join(row_parts[1:])
                if result_dict['리픽싱내용'] == ['-']: result_dict['리픽싱내용'] = "-"

        if '교환대상' in keyword_text:
            if len(row_parts) > 2:
                result_dict['대상주식'] = row_parts[2]

        elif '전환에 따라' in keyword_text:
            if len(row_parts) > 2:
                result_dict['대상주식'] = row_parts[2]

        if '옵션에 관한' in keyword_text:
            if len(row_parts) > 1:
                result_dict['옵션사항'] = ' '.join(row_parts[1:])

    try:
        date_format = "%Y-%m-%d"
        date1 = datetime.strptime(result_dict['만기일'], date_format)
        date2 = datetime.strptime(result_dict['납입일'], date_format)
        diff_days = (date1 - date2).days
        result_dict['만기'] = f"{round(diff_days/365.0, 1)}년"
    except:
        result_dict['만기'] = "-"

    if '리픽싱가격' not in result_dict.keys(): result_dict['리픽싱가격'] = "-"
    if '리픽싱내용' not in result_dict.keys(): result_dict['리픽싱내용'] = "-"
    try:
        result_dict['리픽싱가격'] = f"{round(100 * result_dict['리픽싱가격'] / result_dict['전환가액(원)'],0)}%"
    except:
        result_dict['리픽싱가격'] = "-"
    
    try:
        result_dict['발행대상'], result_dict['검산'] = list_fund_participants(tables)
    except Exception:
        result_dict['발행대상'], result_dict['검산'] = "-", 0.0

    return result_dict

SAMPLE_COMPANY_COUNT = 9999  # high cap; actual batching controls memory

def main():
    with open(GROUPED_FILE, "r", encoding="utf-8") as f: grouped_data = json.load(f)
    grouped_data = grouped_data.get("grouped_by_corp_code")

    def save_batch(batch_results, batch_index):
        out_file = f"reports_details_part_{batch_index}.json"
        with open(out_file, "w", encoding="utf-8") as f:
            json.dump(batch_results, f, ensure_ascii=False, indent=2)
        print(f"Saved {len(batch_results)} reports to {out_file}")

    batch_results = []
    batch_index = 1
    company_counter = 0

    for _, corp_data in grouped_data.items():
        if company_counter >= SAMPLE_COMPANY_COUNT: break

        if company_counter % 10 == 0:
            print(f"Processing company {company_counter}")

        corp_name = corp_data.get("corp_name") 
        stock_code = "A" + corp_data.get("stock_code")

        corp_cls = corp_data.get("corp_cls")
        if corp_cls == "Y": corp_cls = "코스피"
        elif corp_cls == "K": corp_cls = "코스닥"
        else: corp_cls = "오류"

        reports = corp_data.get("reports")
        for report in reports:
            rcept_no = report.get("rcept_no")

            tables = unpack(rcept_no)
            table_data = extract_table_data(report, tables)

            batch_results.append(
                {
                    "corp_name": corp_name,
                    "stock_code": stock_code,
                    "corp_cls": corp_cls,
                    "rcept_no": rcept_no,
                    "table_data": table_data,
                }
            )
        company_counter += 1

        if company_counter % BATCH_SIZE == 0:
            save_batch(batch_results, batch_index)
            batch_results = []
            batch_index += 1

    if batch_results:
        save_batch(batch_results, batch_index)

if __name__ == "__main__":
    main()
