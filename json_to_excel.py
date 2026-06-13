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
                        
                        # 🔥 [버그 수정] Katana JSONL은 request 구조가 아니라 최상위에 데이터가 바로 존재합니다.
                        url_str = obj.get('url', '')
                        if not url_str:
                            continue # URL이 없는 무효 라인은 패스
                            
                        # 타겟 도메인 분리 (예: https://apple.test.com/path -> apple.test.com)
                        if '//' in url_str:
                            target_source = url_str.split('/')[2]
                        else:
                            target_source = url_str.split('/')[0]
                        
                        row = {
                            '대상 타겟 (Target)': target_source,
                            '찾은 시간 (Timestamp)': obj.get('timestamp', ''),
                            '요청 메서드 (Method)': obj.get('method', 'GET'), # 최상위 method 매핑
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
    # 보기 좋게 타겟별 -> URL별로 정렬
    df = df.sort_values(by=['대상 타겟 (Target)', '발견된 URL (URL)'])
else:
    df = pd.DataFrame(columns=['대상 타겟 (Target)', '찾은 시간 (Timestamp)', '요청 메서드 (Method)', '발견된 URL (URL)', '출처 페이지 (Source)', 'HTML 태그 (Tag)', '속성 (Attribute)'])
    df.loc[0] = ['스캔된 결과가 없거나 차단되었습니다.', '', '', '', '', '', '']

# 엑셀 파일 작성
with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='통합 병렬스캔 결과')
    worksheet = writer.sheets['통합 병렬스캔 결과']
    
    # 엑셀 가독성을 위한 열 너비 자동 조절
    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

print(f"🟢 총 {len(data)}개의 데이터 통합 및 엑셀 변환 성공!")
