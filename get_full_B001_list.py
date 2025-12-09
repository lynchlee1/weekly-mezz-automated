import json
import requests
from config import API_KEY

'''
### Purpose ###
    OPENDART API를 이용해 "주요사항보고 > 주요사항보고서" 기존공시 및 정정공시 전부 수집하는 코드
'''

BASE_URL = "https://opendart.fss.or.kr/api/list.json"

QUARTERS = [
    ("Q1", "0101", "0331"),  # 01/01 - 03/31
    ("Q2", "0401", "0630"),  # 04/01 - 06/30
    ("Q3", "0701", "0930"),  # 07/01 - 09/30
    ("Q4", "1001", "1231"),  # 10/01 - 12/31
]

START_YEAR = 2020
END_YEAR = 2025

def collect_reports_for_period(year, quarter_name, bgn_de, end_de):
    """Collect all reports for a specific period with pagination."""
    params = {
        "crtfc_key": API_KEY,       # OPENDART API Key
        "bgn_de": bgn_de,           # 시작일
        "end_de": end_de,           # 종료일
        "last_reprt_at": "N",       # 기존공시 및 정정공시 전부 수집
        "pblntf_ty": "B",           # 주요사항보고
        "pblntf_detail_ty": "B001", # 주요사항보고서
        "page_count": 100,          # 요청당 최대 개수인 100개 수집
    }
    
    # First request to get total page count
    first_response = requests.get(BASE_URL, params=params)
    first_data = first_response.json()
    
    if first_data.get("status") != "000":
        print(f"Error for {year} {quarter_name}: {first_data.get('message', 'API error')}")
        return None
    
    total_page = first_data.get("total_page", 1)
    
    # Collect all reports from all pages
    all_reports = []
    for page_no in range(1, total_page + 1):
        params["page_no"] = page_no
        response = requests.get(BASE_URL, params=params)
        data = response.json()
        
        if data.get("status") == "000" and "list" in data:
            all_reports.extend(data["list"])
            print(f"  Page {page_no}/{total_page}: Collected {len(data['list'])} reports")
    
    # Combine all reports into final response structure
    final_data = {
        "status": "000",
        "message": "정상",
        "year": year,
        "quarter": quarter_name,
        "bgn_de": bgn_de,
        "end_de": end_de,
        "total_count": len(all_reports),
        "total_page": total_page,
        "list": all_reports
    }
    
    return final_data

# Iterate through all years and quarters
for year in range(START_YEAR, END_YEAR + 1):
    print(f"\n{'='*60}")
    print(f"Processing year {year}")
    print(f"{'='*60}")
    
    for quarter_name, start_date, end_date in QUARTERS:
        bgn_de = f"{year}{start_date}"
        end_de = f"{year}{end_date}"
        
        print(f"\nCollecting {year} {quarter_name} ({bgn_de} - {end_de})...")
        
        data = collect_reports_for_period(year, quarter_name, bgn_de, end_de)
        
        if data:
            filename = f"response_{year}_{quarter_name}.json"
            with open(filename, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            
            print(f"  ✓ Saved {len(data['list'])} reports to {filename}")
        else:
            print(f"  ✗ Failed to collect data for {year} {quarter_name}")

print(f"\n{'='*60}")
print("All data collection completed!")
print(f"{'='*60}")
