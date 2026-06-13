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
                line_str = line.strip()
                if not line_str:
                    continue
                    
                url_str = ""
                timestamp = ""
                method = "GET"
                source = ""
                tag = ""
                attribute = ""
                
                # 1. JSON 구조 파싱 시도
                try:
                    obj = json.loads(line_str)
                    
                    # 🔥 [버그 해결] 내포 구조(Standard Katana)와 평면 구조 둘 다 완벽 대응
                    if 'request' in obj and isinstance(obj['request'], dict):
                        url_str = obj['request'].get('url', '')
                        method = obj['request'].get('method', 'GET')
                    else:
                        url_str = obj.get('url', '')
                        method = obj.get('method', 'GET')
                        
                    timestamp = obj.get('timestamp', '')
                    source = obj.get('source', '')
                    tag = obj.get('tag', '')
                    attribute = obj.get('attribute', '')
                    
                except Exception:
                    # 2. 만약 예외 상황으로 일반 텍스트(URL 리스트) 형태로 저장되었을 경우 우회 대응
                    if line_str.startswith('http://') or line_str.startswith('https://'):
                        url_str = line_str

                # URL 정보가 끝까지 안 찾아지면 스킵
                if not url_str:
                    continue
                    
                # 타겟 도메인 분리 (예: https://apple.test.com/path -> apple.test.com)
                try:
                    if '//' in url_str:
                        target_source = url_str.split('/')[2]
                    else:
                        target_source = url_str.split('/')[0]
                except Exception:
                    target_source = "Unknown"
                    
                row = {
                    '대상 타겟 (Target)': target_source,
                    '찾은 시간 (Timestamp)': timestamp,
                    '요청 메서드 (Method)': method,
                    '발견된 URL (URL)': url_str,
                    '출처 페이지 (Source)': source,
                    'HTML 태그 (Tag)': tag,
                    '속성 (Attribute)': attribute
                }
                data.append(row)

# 엑셀 변환 및 정렬
if data:
    df = pd.DataFrame(data)
    df = df.sort_values(by=['대상 타겟 (Target)', '발견된 URL (URL)'])
else:
    df = pd.DataFrame(columns=['대상 타겟 (Target)', '찾은 시간 (Timestamp)', '요청 메서드 (Method)', '발견된 URL (URL)', '출처 페이지 (Source)', 'HTML 태그 (Tag)', '속성 (Attribute)'])
    df.loc[0] = ['스캔된 결과가 없거나 차단되었습니다.', '', '', '', '', '', '']

# 최종 엑셀 파일 생성 및 서식 자동 지정
with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='통합 병렬스캔 결과')
    worksheet = writer.sheets['통합 병렬스캔 결과']
    
    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

print(f"🟢 총 {len(data)}개의 데이터가 정상적으로 통합되어 엑셀로 변환되었습니다!")
