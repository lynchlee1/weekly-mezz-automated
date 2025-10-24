import pandas as pd
import re
import sys
import os

from basics import parse_number

def fundname_to_corpname(fundname: str) -> str:
    corpnames = {
        "2": {
            "안다": "안다", "아샘": "아샘", "수성": "수성", "리딩": "리딩", "월넛": "월넛", 
            "이현": "이현", "타임": "타임", 
            "DB": "DB증권", "SP": "SP", "SH": "SH",
        },
        "3": {
            "칸서스": "칸서스", "나이스": "나이스", "포커스": "포커스", "마스터": "마스터",
            "와이씨": "와이씨", "코람코": "코람코", "이지스": "이지스", "시너지": "시너지",
            "라이프": "라이프", "아트만": "아트만", "에이원": "에이원", "레이크": "레이크",
            "문스톤": "문스톤",
            "GVA": "GVA",
        },
        "4": {
            "라이노스": "라이노스", "르네상스": "르네상스", "한국밸류": "한국밸류", "오라이언": "오라이언",
            "피보나치": "피보나치", "마일스톤": "마일스톤", "삼성증권": "삼성증권", "아스트라": "아스트라", 
            "이아이피": "이아이피", "썬앤트리": "썬앤트리", "지베스코": "지베스코", "씨스퀘어": "씨스퀘어",
            "NH앱솔": "NH헤지", "NH헤지": "NH헤지", "디비증권": "DB증권", 
        },
        "5": {
            "지브이에이": "GVA", "IBK투자": "IBK투자", "브로드하이": "브로드하이", "코리아에셋": "코리아에셋",
        },
        "6": {
            "아이트러스트": "아이트러스트",
        },
        "7": {},
        "8": {
            "아이비케이캐피탈": "IBK캐피탈"
        }

    }
    fundname_no_space = fundname.replace(' ', '')
    for i in range(2, len(corpnames)-1):
        prefix = fundname_no_space[:i]
        for match, corp_name in corpnames[str(i)].items():
            if match == prefix: return corp_name
    return fundname_no_space

def list_fund_participants(all_tables, save_path = None): 
    first_table, second_table = None, None
    for idx in range(len(all_tables) - 1, -1, -1): # Search from the end in reverse direction to find the LAST matching table
        table = all_tables[idx]
        table_text = table.get_text()
        if '발행 대상자명' in table_text:
            first_table = table
            if idx + 1 < len(all_tables): second_table = all_tables[idx + 1]
            else: second_table = None
            break

    if first_table is not None and second_table is not None:
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
        
        second_table_rows = []
        tbody = second_table.find('tbody')
        if tbody:
            for tr in tbody.find_all('tr'):
                cells = tr.find_all('td')
                if cells:
                    row = [cell.get_text(strip=True) for cell in cells]
                    second_table_rows.append(row)

        first_table_df = pd.DataFrame(first_table_rows)
        second_table_df = pd.DataFrame(second_table_rows)

        '''
        Now Table of Buyers is set
        '''

        if not first_table_df.empty:
            idx_col1 = -1
            idx_col2 = -1
            for i, col in enumerate(first_table_df.iloc[0, :]):
                if '발행 대상자명' in col: idx_col1 = i
                if '권면' in col or '총액' in col or '금액' in col: idx_col2 = i
            first_table_filtered = first_table_df.iloc[1:, [idx_col1, idx_col2]].copy()
            first_table_filtered.columns = ['발행 대상자명', '권면']
            
            def extract_bonken_numbers(text):
                if pd.isna(text) or not isinstance(text, str): return text

                if '(' in text and '본건' in text: # Check if text contains '(' and '본건'
                    # Handle both cases: "text (본건...)" and "(본건...) text"
                    if text.startswith('('):
                        before_paren = text.split(')')[0] if ')' in text else text
                        numbers = re.findall(r'\d+', before_paren) # find all numbers
                    else:
                        after_paren = text.split('(')[1] if '(' in text else text
                        numbers = re.findall(r'\d+', after_paren) # find all numbers
                    
                    if numbers: return '|'.join(numbers) # Join numbers with '|' separator
                return text  

            first_table_filtered['발행 대상자명'] = first_table_filtered['발행 대상자명'].apply(extract_bonken_numbers)
            first_table_filtered['권면'] = first_table_filtered['권면'].apply(parse_number)
        else: first_table_filtered = pd.DataFrame(columns=['발행 대상자명', '권면'])
        print("first_table_filtered:", first_table_filtered)
        
        if not second_table_df.empty and len(second_table_df.columns) >= 2:
            first_cell = str(second_table_df.iloc[0, 0])
            has_header = not bool(re.search(r'\d', first_cell))            
            if has_header: data_rows = second_table_df.iloc[1:, [0, 1]].copy()
            else: data_rows = second_table_df.iloc[:, [0, 1]].copy()
            data_rows.columns = ['구분', '본건펀드']
            
            def extract_fund_numbers(text):
                if pd.isna(text) or not isinstance(text, str): return text
                if '본건' in text or '펀드' in text:
                    numbers = re.findall(r'\d+', text)
                    if numbers: return numbers[0]
                return text

            data_rows['구분'] = data_rows['구분'].apply(extract_fund_numbers)
            data_rows['본건펀드'] = data_rows['본건펀드'].apply(fundname_to_corpname)
            
            # Group by column 2 (본건펀드) and join the 구분 values
            grouped = data_rows.groupby('본건펀드')['구분'].apply(lambda x: '|'.join(x.astype(str))).reset_index()
            grouped.columns = ['본건펀드', '구분']
            
            # Add third column: sum values from first table based on matching fund numbers
            def sum_from_first_table(fund_numbers_str):
                if pd.isna(fund_numbers_str) or not fund_numbers_str: return 0
                fund_numbers = str(fund_numbers_str).split('|')
                total_sum = 0
                for fund_num in fund_numbers:
                    fund_num = fund_num.strip()
                    if fund_num:
                        matching_rows = first_table_filtered[
                            first_table_filtered['발행 대상자명'].astype(str).str.contains(fund_num, na=False)
                        ]
                        if not matching_rows.empty:
                            total_sum += matching_rows['권면'].sum()
                return total_sum
            
            grouped['합계'] = grouped['구분'].apply(sum_from_first_table)
            
            def is_fund_number_only(text):
                if pd.isna(text) or not isinstance(text, str):
                    return False
                cleaned = re.sub(r'[\d\|\s]', '', str(text))
                return len(cleaned) == 0
            non_fund_rows = first_table_filtered[
                ~first_table_filtered['발행 대상자명'].apply(is_fund_number_only)
            ].copy()
            
            if not non_fund_rows.empty:
                non_fund_rows['구분'] = '-'
                non_fund_rows['본건펀드'] = non_fund_rows['발행 대상자명']
                non_fund_rows['합계'] = non_fund_rows['권면']
                non_fund_rows = non_fund_rows[['구분', '본건펀드', '합계']]
                second_table_filtered = pd.concat([grouped, non_fund_rows], ignore_index=True)
                print("second_table_filtered:", second_table_filtered)
            else:
                second_table_filtered = grouped
                print("second_table_filtered:", second_table_filtered)
            second_table_filtered = second_table_filtered.sort_values('합계', ascending=False).reset_index(drop=True)
        else: 
            if not first_table_filtered.empty:
                def clean_corp_name(name):
                    if pd.isna(name) or not isinstance(name, str):
                        return name
                    cleaned = str(name)
                    cleaned = cleaned.replace('(주)', '').replace('주식회사 ', '').replace(' 주식회사', '').strip()
                    return cleaned
                
                first_table_filtered['구분'] = '-'
                first_table_filtered['본건펀드'] = first_table_filtered['발행 대상자명'].apply(clean_corp_name)
                first_table_filtered['합계'] = first_table_filtered['권면']
                second_table_filtered = first_table_filtered[['본건펀드', '구분', '합계']].copy()
            else:
                second_table_filtered = pd.DataFrame(columns=['본건펀드', '구분', '합계'])

        def text_fund_participants(second_table_filtered):
            parts = []
            for _, row in second_table_filtered.iterrows():
                name = str(row.iloc[0])
                value = row.iloc[2] / 10**8
                if value.is_integer():
                    value_str = f"{int(value)}"
                else:
                    value_str = f"{value:.1f}"
                parts.append(f"{name} {value_str}")
            return ', '.join(parts)

        if save_path is not None:
            if getattr(sys, 'frozen', False): exe_dir = os.path.dirname(sys.executable)
            else: exe_dir = os.path.dirname(os.path.abspath(__file__))
            output_path = os.path.join(exe_dir, save_path)
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer: second_table_filtered.to_excel(writer, sheet_name='Final Table', index=False)
            return f"Saved results to {output_path}"
        else: 
            return text_fund_participants(second_table_filtered)

if __name__ == "__main__":
    from main import unpack
    all_tables = unpack("20251016000315")
    print(list_fund_participants(all_tables))