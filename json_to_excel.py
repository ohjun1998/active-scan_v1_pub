# json_to_excel.py
import json
import os
import glob
import pandas as pd

output_dir = 'katana_outputs'
excel_file = 'katana_multi_scan_report.xlsx'
data = []

# 각 VM에서 수집되어 복호화된 jsonl 파일들 매핑
jsonl_files = glob.glob(os.path.join(output_dir, 'part_*.jsonl'))

if jsonl_files:
    for file_path in jsonl_files:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                if line.strip():
                    try:
                        obj = json.loads(line)
                        # 어떤 타겟에서 나온 결과인지 수집 (도메인 분리)
                        url_str = obj.get('request', {}).get('url', '')
                        target_source = url_str.split('/')[2] if '/' in url_str else 'Unknown'
                        
                        row = {
                            '대상 타겟 (Target)': target_source,
                            '찾은 시간 (Timestamp)': obj.get('timestamp', ''),
                            '요청 메서드 (Method)': obj.get('request', {}).get('method', 'GET'),
                            '발견된 URL (URL)': url_str,
                            '출처 페이지 (Source)': obj.get('source', ''),
                            'HTML 태그 (Tag)': obj.get('tag', ''),
                            '속성 (Attribute)': obj.get('attribute', '')
                        }
                        data.append(row)
                    except Exception as e:
                        pass

# 엑셀 변환 및 정렬
if data:
    df = pd.DataFrame(data)
    df = df.sort_values(by=['대상 타겟 (Target)', '발견된 URL (URL)'])
else:
    df = pd.DataFrame(columns=['대상 타겟 (Target)', '찾은 시간', '메서드', 'URL', '출처', '태그', '속성'])
    df.loc[0] = ['스캔된 결과가 없거나 전부 차단되었습니다.', '', '', '', '', '', '']

with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='통합 병렬스캔 결과')
    worksheet = writer.sheets['통합 병렬스캔 결과']
    
    # 엑셀 가독성을 위한 열 너비 자동 조절
    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

print("🟢 모든 가상머신 데이터 통합 및 엑셀 변환 성공!")
