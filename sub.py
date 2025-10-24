import pandas as pd
import re
import sys
import os

from basics import parse_number

CORPNAMES = {
    "2": {
        "안다": "안다", "아샘": "아샘", "수성": "수성", "리딩": "리딩", "월넛": "월넛", 
        "이현": "이현", "타임": "타임", "신한": "신한",
        "SP": "SP", "SH": "SH",
    },
    "3": {
        "칸서스": "칸서스", "나이스": "나이스", "포커스": "포커스", "마스터": "마스터",
        "와이씨": "와이씨", "코람코": "코람코", "이지스": "이지스", "시너지": "시너지",
        "라이프": "라이프", "아트만": "아트만", "에이원": "에이원", "레이크": "레이크",
        "문스톤": "문스톤", "크로톤": "크로톤",
        "GVA": "GVA",
    },
    "4": {
        "라이노스": "라이노스", "르네상스": "르네상스", "한국밸류": "한국밸류", "오라이언": "오라이언",
        "피보나치": "피보나치", "마일스톤": "마일스톤", "삼성증권": "삼성증권", "아스트라": "아스트라", 
        "이아이피": "이아이피", "썬앤트리": "썬앤트리", "지베스코": "지베스코", "씨스퀘어": "씨스퀘어",
        "인피니티": "인피니티", "코너스톤": "코너스톤", "키움증권": "키움증권",
        "NH앱솔": "NH헤지", "NH헤지": "NH헤지", "디비증권": "DB증권", "DB증권": "DB증권",
    },
    "5": {
        "지브이에이": "GVA", "IBK투자": "IBK투자", "브로드하이": "브로드하이", "코리아에셋": "코리아에셋",
    },
    "6": {
        "아이트러스트": "아이트러스트", "엔에이치투자": "NH투자",
    },
    "7": {},
    "8": {
        "아이비케이캐피탈": "IBK캐피탈", "제이비우리캐피탈": "JB우리캐피탈",
    }
}

def preprocess_fundname(fundname: str) -> str:
    replacements = ["주식회사 ", " 주식회사", "(주)"]
    for replacement in replacements:
        fundname = fundname.replace(replacement, '')
    return fundname

def fundname_to_corpname(fundname: str) -> str:
    corpnames = CORPNAMES
    fundname = preprocess_fundname(fundname)
    fundname_no_space = fundname.replace(' ', '')
    for i in range(2, len(corpnames)-1):
        prefix = fundname_no_space[:i]
        for match, corp_name in corpnames[str(i)].items():
            if match == prefix: return corp_name
    return fundname_no_space

def fundname_to_corpname_safe(fundname: str) -> str:
    matches = []
    fundname = preprocess_fundname(fundname)
    fundname_no_space = fundname.replace(' ', '')
    corpnames = CORPNAMES
    for _, corp_dict in corpnames.items():
        for match, corp_name in corp_dict.items():
            if match in fundname_no_space:
                matches.append(corp_name)
    if len(matches) > 1:
        return "-".join(matches)
    elif len(matches) == 1:
        return matches[0]
    return fundname_no_space

def list_fund_participants(all_tables): 
    first_table, second_table = None, None
    for idx in range(len(all_tables) - 1, -1, -1): # Search from the end in reverse direction to find the LAST matching table
        table = all_tables[idx]
        table_text = table.get_text()
        if '발행 대상자명' in table_text:
            first_table = table
            if idx + 1 < len(all_tables): second_table = all_tables[idx + 1]
            else: second_table = None
            break

    if first_table is not None:
        first_table_rows = []
        thead = first_table.find('thead')
        if thead:
            for tr in thead.find_all('tr'):
                cells = tr.find_all('th')
                if cells:
                    header_row = [cell.get_text(strip=True) for cell in cells]
                    first_table_rows.append(header_row)
        
        tbody = first_table.find('tbody')
        if tbody:
            for tr in tbody.find_all('tr'):
                all_cells = tr.find_all('te')
                if all_cells:
                    row = [cell.get_text(strip=True) for cell in all_cells]
                    first_table_rows.append(row)
        first_table_df = pd.DataFrame(first_table_rows)

        if not first_table_df.empty and len(first_table_df.columns) >= 2:
            idx_col1 = -1
            idx_col2 = -1
            for i, col in enumerate(first_table_df.iloc[0, :]):
                if '발행 대상자명' in col: idx_col1 = i
                if '권면' in col or '총액' in col or '금액' in col: idx_col2 = i
            first_table_filtered = first_table_df.iloc[1:, [idx_col1, idx_col2]].copy()
            first_table_filtered.columns = ['발행 대상자명', '권면']
            
            def extract_bonken_numbers(text):
                if '(' in text and '본건' in text: # Check if text contains '(' and '본건'
                    # Handle both cases: "corpname+(본건#)" and "(본건#)+corpname"
                    if text.startswith('('):
                        before_paren = text.split(')')[0] if ')' in text else text
                        numbers = re.findall(r'\d+', before_paren) # find all numbers
                    else:
                        after_paren = text.split('(')[1] if '(' in text else text
                        numbers = re.findall(r'\d+', after_paren) # find all numbers
                    
                    if numbers: return '|'.join(numbers) # Join numbers with '|' separator
                return text  

            first_table_filtered['발행 대상자명'] = first_table_filtered['발행 대상자명'].apply(extract_bonken_numbers)
            first_table_filtered['발행 대상자명'] = first_table_filtered['발행 대상자명'].apply(fundname_to_corpname_safe)
            first_table_filtered['권면'] = first_table_filtered['권면'].apply(parse_number)
        else: first_table_filtered = pd.DataFrame(columns=['발행 대상자명', '권면'])
    
    if second_table is not None:
        second_table_rows = []
        tbody = second_table.find('tbody')
        if tbody:
            for tr in tbody.find_all('tr'):
                cells = tr.find_all('td')
                if cells:
                    row = [cell.get_text(strip=True) for cell in cells]
                    second_table_rows.append(row)
        second_table_df = pd.DataFrame(second_table_rows)

        if not second_table_df.empty and len(second_table_df.columns) >= 2:
            first_cell = str(second_table_df.iloc[0, 0])
            has_header = not bool(re.search(r'\d', first_cell))            
            if has_header:
                second_table_filtered = second_table_df.iloc[1:, [0, 1]].copy()
            else: 
                second_table_filtered = second_table_df.iloc[:, [0, 1]].copy()
            second_table_filtered.columns = ['구분', '본건펀드']

            def extract_fund_numbers(text):
                if '본건' in text or '펀드' in text:
                    numbers = re.findall(r'\d+', text)
                    if numbers: return numbers[0]
                return text

            second_table_filtered['구분'] = second_table_filtered['구분'].apply(extract_fund_numbers)
            second_table_filtered['본건펀드'] = second_table_filtered['본건펀드'].apply(fundname_to_corpname)

        '''
        First Table : corpname(safe) or bonken numbers | fiscal amount
        Second Table : bonken number | corpname
        '''

    if first_table is not None and second_table is not None:
        def map_numbers_to_corpnames(text):
            if '|' in text:
                first_num = text.split('|')[0].strip()
                if first_num.isdigit():
                    if not second_table_df.empty and len(second_table_df.columns) >= 2:
                        matching_rows = second_table_df[second_table_df.iloc[:, 0].astype(str).str.contains(first_num, na=False)]
                        if not matching_rows.empty:
                            corp_name = matching_rows.iloc[0, 1]
                            return fundname_to_corpname(str(corp_name))
                return text
            else:
                # Handle single number
                if text.isdigit():
                    if not second_table_df.empty and len(second_table_df.columns) >= 2:
                        matching_rows = second_table_df[second_table_df.iloc[:, 0].astype(str).str.contains(text, na=False)]
                        if not matching_rows.empty:
                            corp_name = matching_rows.iloc[0, 1]
                            return fundname_to_corpname(str(corp_name))
                return text
        
        if not first_table_filtered.empty:
            first_table_filtered['발행 대상자명'] = first_table_filtered['발행 대상자명'].apply(map_numbers_to_corpnames)
            final_table = first_table_filtered.groupby('발행 대상자명')['권면'].sum().reset_index()
            final_table = final_table.sort_values('권면', ascending=False).reset_index(drop=True)
            total_amount = final_table['권면'].sum() / 10**8

            def format_final_table_text(final_table):
                parts = []
                for _, row in final_table.iterrows():
                    corp_name = str(row['발행 대상자명'])
                    value = row['권면'] / 10**8
                    if value.is_integer():
                        value_str = f"{int(value)}"
                    else:
                        value_str = f"{value:.1f}"
                    parts.append(f"{corp_name} {value_str}")
                return ', '.join(parts)
            
            result_text = format_final_table_text(final_table)
            return result_text, total_amount
    
    elif first_table is not None:
        total_amount = first_table_filtered['권면'].sum() / 10**8
        return format_final_table_text(first_table_filtered), total_amount
        
    return "-", 0.0

if __name__ == "__main__":
    from main import unpack
    all_tables = unpack("20250801000380")
    print(list_fund_participants(all_tables))
