import json
import os
from pathlib import Path

RESPONSES_FOLDER = "responses"
EXCLUSION_KEYWORDS = ["[첨부정정]", "[첨부추가]"]
INCLUSION_KEYWORDS = [["전환사채", "교환사채", "신주인수권부사채"], "발행"]

def should_include_report(report):
    corp_cls = report.get("corp_cls", "")
    if corp_cls not in ["Y", "K"]: return False
    
    report_nm = report.get("report_nm", "")
    for keyword in EXCLUSION_KEYWORDS:
        if keyword in report_nm: return False

    for keyword_group in INCLUSION_KEYWORDS:
        group_matched = False
        if isinstance(keyword_group, list):
            for keyword in keyword_group:
                if keyword in report_nm:
                    group_matched = True
                    break
        else:
            if keyword_group in report_nm:
                group_matched = True
        
        if not group_matched:
            return False
    
    return True

def process_all_json_files():
    responses_path = Path(RESPONSES_FOLDER)
    json_files = sorted(responses_path.glob("response_*.json"))

    all_filtered_reports = []
    total_original = 0
    total_filtered = 0
    for json_file in json_files:
        with open(json_file, "r", encoding="utf-8") as f: data = json.load(f)            
        reports = data.get("list", [])
        original_count = len(reports)
        total_original += original_count
        filtered_reports = [report for report in reports if should_include_report(report)]
        filtered_count = len(filtered_reports)
        total_filtered += filtered_count
        all_filtered_reports.extend(filtered_reports)
    
    output_data = {
        "total_original_count": total_original,
        "total_filtered_count": total_filtered,
        "list": all_filtered_reports
    }
    output_filename = "filtered_B001_list.json"
    with open(output_filename, "w", encoding="utf-8") as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)
    print(f"Processing completed! Results saved to: {output_filename}")

    unique_titles = set()
    for report in all_filtered_reports:
        title = report.get("report_nm")
        if title:
            unique_titles.add(title)
    print("Unique report titles (in filtered set):")
    for t in sorted(unique_titles): print(" -", t)

def group_by_corp_code():
    with open("filtered_B001_list.json", "r", encoding="utf-8") as f: data = json.load(f)
    reports = data.get("list", [])
    
    grouped = {}
    for report in reports:
        corp_code = report.get("corp_code")
        if corp_code not in grouped:
            grouped[corp_code] = []
        grouped[corp_code].append(report)
    
    grouped_data = {}
    for corp_code, company_reports in grouped.items():
        latest_report = max(company_reports, key=lambda x: x.get("rcept_dt", ""))
        company_info = {
            "corp_code": latest_report.get("corp_code"),
            "corp_name": latest_report.get("corp_name"),
            "stock_code": latest_report.get("stock_code"),
            "corp_cls": latest_report.get("corp_cls")
        }
        
        cleaned_reports = []
        for report in company_reports:
            cleaned_report = {
                "report_nm": report.get("report_nm"),
                "rcept_no": report.get("rcept_no"),
                "rcept_dt": report.get("rcept_dt"),
                "rm": report.get("rm")
            }
            cleaned_reports.append(cleaned_report)
        
        grouped_data[corp_code] = {
            **company_info,
            "reports": cleaned_reports
        }
    
    output_data = {
        "total_companies": len(grouped_data),
        "total_reports": len(reports),
        "grouped_by_corp_code": grouped_data
    }
    
    output_filename = "filtered_B001_list_grouped.json"
    with open(output_filename, "w", encoding="utf-8") as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    # process_all_json_files()
    group_by_corp_code()

