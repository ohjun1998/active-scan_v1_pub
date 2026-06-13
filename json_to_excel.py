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

# 🚀 [추적 모니터링 로그] 파이썬이 실제로 어떤 파일들을 찾았는지 출력해 줍니다.
print(f"🔍 [파이썬 디버그] 탐색된 수집 파일 목록: {jsonl_files}")

if jsonl_files:
    for file_path in jsonl_files:
        print(f"📖 현재 파일 분석 중: {file_path}")
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
                
                try:
                    obj = json.loads(line_str)
                    
                    # 구조 매핑 완벽 대응
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
                    if line_str.startswith('http://') or line_str.startswith('https://'):
                        url_str = line_str

                if not url_str:
                    continue
                    
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
    print(f"🟢 [성공] 총 {len(data)}개의 가용한 엔드포인트를 매핑했습니다.")
else:
    df = pd.DataFrame(columns=['대상 타겟 (Target)', '찾은 시간 (Timestamp)', '요청 메서드 (Method)', '발견된 URL (URL)', '출처 페이지 (Source)', 'HTML 태그 (Tag)', '속성 (Attribute)'])
    df.loc[0] = ['스캔된 결과가 없거나 차단되었습니다.', '', '', '', '', '', '']
    print("⚠️ [경고] 처리할 데이터가 하나도 없습니다. 데이터 풀을 확인하세요.")

with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='통합 병렬스캔 결과')
    worksheet = writer.sheets['통합 병렬스캔 결과']
    
    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
