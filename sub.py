import pandas as pd
import re
import sys
import os

from basics import parse_number

CORPNAMES = {
    "1": {

    },
    "2": {
        "안다": "안다", "아샘": "아샘", "수성": "수성", "리딩": "리딩", "월넛": "월넛", 
        "이현": "이현", "타임": "타임", "삼성": "삼성", "키움": "키움", "노앤": "노앤",
        "코어": "코어", "한양": "한양", "교보": "교보", "현대": "현대", "람다": "람다",
        "아름": "아름", "컴파": "컴파", "대신": "대신", "킹고": "킹고", "신한": "신한", 
        "이앤": "이앤VC", "흥국": "흥국증권",
        "SP": "SP", "SH": "SH", "NH": "NH", "KY": "KY", "JW": "JW",
    },
    "3": {
        "칸서스": "칸서스", "나이스": "나이스", "포커스": "포커스", "마스터": "마스터",
        "와이씨": "와이씨", "코람코": "코람코", "이지스": "이지스", "시너지": "시너지",
        "라이프": "라이프", "아트만": "아트만", "에이원": "에이원", "레이크": "레이크",
        "문스톤": "문스톤", "크로톤": "크로톤", "프렌드": "프렌드", "파로스": "파로스",
        "라이언": "라이언", "린드먼": "린드먼", "블랙펄": "블랙펄", "에이스": "에이스",
        "유암코": "유암코", "타이거": "타이거", "스마일": "스마일", "디파인": "디파인",
        "가이아": "가이아",
        "GVA": "GVA", "제이비": "JB", "안다H": "안다H", "케이비": "KB"
    },
    "4": {
        "라이노스": "라이노스", "르네상스": "르네상스", "한국밸류": "한국밸류", "오라이언": "오라이언",
        "피보나치": "피보나치", "마일스톤": "마일스톤", "삼성증권": "삼성증권", "아스트라": "아스트라", 
        "이아이피": "이아이피", "썬앤트리": "썬앤트리", "지베스코": "지베스코", "씨스퀘어": "씨스퀘어",
        "인피니티": "인피니티", "코너스톤": "코너스톤", "키움증권": "키움증권", "히스토리": "히스토리",
        "한화투자": "한화증권", "신한자산": "신한자산", "신한투자": "신한증권", "다올투자": "다올증권",
        "에이치알": "에이치알", "셀레니언": "셀레니언", "한국투자": "한국투자", "파라투스": "파라투스",
        "윈베스트": "윈베스트", "파인밸류": "파인밸류", "트러스톤": "트러스톤", "패스웨이": "패스웨이",
        "NH앱솔": "NH헤지", "NH헤지": "NH헤지", "디비증권": "DB증권", "DB증권": "DB증권",
        "NH투자": "NH증권", "엔에이치": "NH",
        "KDBC": "KDBC", "IBKC": "IBKC", "디에스씨": "DSC인베",
    },
    "5": {
        "지브이에이": "GVA", "IBK투자": "IBK증권", "브로드하이": "브로드하이", "코리아에셋": "코리아에셋",
        "알파플러스": "알파플러스", "아이비케이": "IBK", "컴퍼니케이": "컴퍼니케이",
        "케이비증권": "KB증권", "제이씨에셋": "JC자산", "안다에이치": "안다H",
    },
    "6": {
        "아이트러스트": "아이트러스트","미래에셋증권": "미래에셋증권", "하나에스앤비": "하나에스앤비", 
        "유진투자증권": "유진증권", 
        "하이투자증권": "IM증권", "엔에이치투자": "NH증권", "대신그로쓰캡": "대신PE",
        "아이비케이씨": "IBKC", "한국투자증권": "한투", "비엔케이증권": "BNK증권",
    },
    "7": {
        "아이비케이투자": "IBK증권", "디에스투자증권": "DS증권",
    },
    "8": {
        "이베스트투자증권": "이베스트증권", "비엔케이투자증권": "비엔케이증권",
        "아이비케이캐피탈": "IBK캐피탈", "제이비우리캐피탈": "JB우리캐피탈",
    },
    "9": {
        "IPARTNERS": "아이파트너스",
    },
    "10": {
        "HYUNSTEADY": "HYUNSTEADY",
    }
}

def preprocess_fundname(fundname: str) -> str:
    replacements = ["주식회사 ", " 주식회사", "(주)"]
    for replacement in replacements:
        fundname = fundname.replace(replacement, '')
    return fundname

def fundname_to_corpname(fundname: str) -> str:
    corpnames = CORPNAMES
    fundname = fundname.replace(' ','')
    replacements = [' ', '주식회사', '(주)', '㈜']
    for replacement in replacements:
        fundname = fundname.replace(replacement, '')

    is_shingisa = False
    if '신기술' in fundname and '조합' in fundname: is_shingisa = True

    corp_found = ""
    for i in range(len(corpnames), 0, -1): 
        prefix = fundname[:i]
        remaining = fundname[i:]
        for match, corp_name in corpnames[str(i)].items():
            if match == prefix: 
                corp_found = corp_name

        found_names = []
        if corp_found: 
            found_names = [corp_found]
            for i in range(len(corpnames), 0, -1):
                for match, corp_name in corpnames[str(i)].items():
                    if match in remaining: 
                        found_names.append(corp_name)

        if corp_found:
            break

    if found_names:
        corp_found = '-'.join(found_names)
    
    if corp_found and is_shingisa: fundname = corp_found + " 신기사"
    if corp_found and not is_shingisa: fundname = corp_found

    return fundname

def fundname_to_corpname_safe(fundname: str) -> str:
    if "-" in fundname: return preprocess_fundname(fundname)
    else: return fundname_to_corpname(fundname)

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

        if not first_table_df.empty and len(first_table_df.columns) >= 2 and len(first_table_df) > 0:
            idx_col1 = -1
            idx_col2 = -1
            for i, col in enumerate(first_table_df.iloc[0, :]):
                if '발행 대상자명' in col: idx_col1 = i
                if '권면' in col or '총액' in col or '금액' in col: idx_col2 = i
            if idx_col1 != -1 and idx_col2 != -1:
                first_table_filtered = first_table_df.iloc[1:, [idx_col1, idx_col2]].copy()
            else:
                first_table_filtered = pd.DataFrame(columns=['발행 대상자명', '권면'])
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
            first_table_filtered['발행 대상자명'] = first_table_filtered['발행 대상자명'].apply(fundname_to_corpname)
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

        if not second_table_df.empty and len(second_table_df.columns) >= 2 and len(second_table_df) > 0:
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
                    if not second_table_df.empty and len(second_table_df.columns) >= 2 and len(second_table_df) > 0:
                        matching_rows = second_table_df[second_table_df.iloc[:, 0].astype(str).str.contains(first_num, na=False)]
                        if not matching_rows.empty and len(matching_rows) > 0:
                            corp_name = matching_rows.iloc[0, 1]
                            return fundname_to_corpname(str(corp_name))
                return text
            else:
                # Handle single number
                if text.isdigit():
                    if not second_table_df.empty and len(second_table_df.columns) >= 2 and len(second_table_df) > 0:
                        matching_rows = second_table_df[second_table_df.iloc[:, 0].astype(str).str.contains(text, na=False)]
                        if not matching_rows.empty and len(matching_rows) > 0:
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

test_cases = [
    '20220706000148', '20231027000166', '20240216000966', '20230508000614', 
    '20240604000386', '20240726000500', '20230807000401', '20230829000575', '20230920000049', 
    '20231027000378', '20231115000399', '20240117000341', '20240125000601', '20240329002828', 
    '20240402003067', '20250409001971', '20240605000268', '20240628000114', '20240827000618', 
    '20250827000516', '20240909000134', '20250916000305', '20250917000329', '20241016000313', 
    '20250912000464', '20250912000257', '20250915000241', '20240425000614', '20250922000288', 
    '20250219001783', '20250724000361', '20250729000174', '20250801000380', '20250807000166', 
    '20250905000042', '20250902000329', '20250908000110', '20250922000174', '20251015000373', 
    '20251001000656', '20250919000150'
]
if __name__ == "__main__":
    from main import unpack
    for test_case in test_cases[:5]:
        all_tables = unpack(test_case)
        result, total_amount = list_fund_participants(all_tables)
        print(f"{test_case} : {result}, {total_amount}")
