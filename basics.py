import requests
import json
import sys
import os
import re
from datetime import datetime
from bs4 import BeautifulSoup
import zipfile
from io import BytesIO

__all__ = ['resource_dir', 'get_api_key', 'split', 'parse_number', 'parse_date', 'get_reports_range', 'unpack']

def resource_dir() -> str:
    """Return the directory where resources (like config.json) live.
    In a frozen EXE, use the executable's directory; otherwise use this file's directory.
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def get_api_key():
    """Get the current API key directly from JSON file in resource directory"""
    config_path = os.path.join(resource_dir(), 'config.json')
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
            return config.get('API_KEY', '')
    except (FileNotFoundError, json.JSONDecodeError, KeyError) as e:
        print(f"Error loading API key from JSON: {e}")
        return ""

def split(text: str) -> list:
    '''
    input: text
    output: list of texts that is splitted by '|'
    '''
    return [part.strip() for part in text.split('|')]

def parse_number(text: str) -> float:
    '''
    input: text in string format. ',' is allowed
    output: number
    '''
    try:
        return float(text.replace(',', ''))
    except Exception:
        return -1.0 # returns float to prevent breaking whole process

def parse_date(text: str) -> str:
    '''
    input: text in string format. 'YYYY-MM-DD', 'YYYY.MM.DD', or 'YYYY년 MM월 DD일' is allowed
    output: date in string format %Y-%m-%d
    '''
    try:
        if re.match(r'^\d{4}\.\d{2}\.\d{2}$', text):
            text = text.replace('.', '-')
        elif re.match(r'^\d{4}년\s?\d{1,2}월\s?\d{1,2}일$', text):
            parts = re.findall(r'\d+', text)
            y, m, d = parts[0], parts[1].zfill(2), parts[2].zfill(2)
            text = f"{y}-{m}-{d}"
        return datetime.strptime(text, '%Y-%m-%d').strftime('%Y-%m-%d')
    except:
        return "-"

def unpack(rcept_no: str) -> list:
    '''
    input: rcept_no
    output: list of tables in the report
    '''
    source_download_url = "https://opendart.fss.or.kr/api/document.xml"
    url = f"{source_download_url}?crtfc_key={get_api_key()}&rcept_no={rcept_no}"
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
                        print(f"Error decoding with utf-8: {name}")
                        try: text = data.decode('cp949')
                        except UnicodeDecodeError: 
                            print(f"Error decoding with cp949: {name}")
                            text = None

                    if text is not None:
                        soup = BeautifulSoup(text, 'html.parser')
                        tables = soup.find_all('table')
                        all_tables.extend(tables)
    except zipfile.BadZipFile: pass
    return all_tables

def get_reports_range(start_date, end_date):
    '''
    input: start_date and end_date in string format %Y%m%d
    output: list of reports between start_date and end_date
    '''
    base_url = "https://opendart.fss.or.kr/api/list.json"
    base_params = {
        'crtfc_key': get_api_key(),
        'bgn_de': start_date,
        'end_de': end_date,
        'last_reprt_at': 'Y',
        'pblntf_ty': 'B', # Additional B001 type is sometimes faulty 
        'page_count': "100"
    }

    results = []
    page_no = 1
    while True: # Iterate through all pages
        params = base_params.copy()
        params['page_no'] = page_no
        response = requests.get(base_url, params=params)
        if response.json()['status'] != '000': break
        results.extend(response.json()['list'])
        if len(response.json()['list']) >= 100: 
            page_no += 1
        else: break
    return results

