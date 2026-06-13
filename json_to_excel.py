# json_to_excel.py
import json
import os
import glob
import re
import pandas as pd

output_dir = 'katana_outputs'
excel_file = 'katana_multi_scan_report.xlsx'
data = []

# 🔥 [핵심 추가] Secrets의 TARGET_URLS 환경변수를 읽어와 허용된 도메인 세트 생성
target_urls_env = os.environ.get('TARGET_URLS', '')
allowed_domains = set()

for line in target_urls_env.split('\n'):
    line_str = line.strip()
    if line_str:
        # 프로토콜(http://, https://) 제거
        domain = re.sub(r'^https?://', '', line_str)
        # 하위 경로(/...) 제거하여 순수 도메인만 추출 (예: obank.kbstar.com)
        domain = domain.split('/')[0].split(':')[0]
        allowed_domains.add(domain)

print(f"🎯 [필터 켜짐] 다음 지정 도메인 결과만 엑셀에 저장됩니다: {allowed_domains}")

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
                
                try:
                    obj = json.loads(line_str)
                    if isinstance(obj, dict):
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
                    pass

                if not url_str:
                    url_match = re.search(r'(https?://[^\s"\'},]+)', line_str)
                    if url_match:
                        url_str = url_match.group(1)

                if not url_str:
                    continue
                    
                # 발견된 URL의 도메인 추출
                try:
                    if '//' in url_str:
                        target_source = url_str.split('/')[2].split(':')[0]
                    else:
                        target_source = url_str.split('/')[0].split(':')[0]
                except Exception:
                    target_source = "Unknown"
                    
                # 🔥 [화이트리스트 필터링] 허용 도메인 목록에 없으면 엑셀에 넣지 않고 스킵!
                if target_source not in allowed_domains:
                    continue
                    
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
    print(f"🟢 [필터링 완료] 총 {len(data)}개의 순정 타겟 엔드포인트가 엑셀에 매핑되었습니다!")
else:
    df = pd.DataFrame(columns=['대상 타겟 (Target)', '찾은 시간 (Timestamp)', '요청 메서드 (Method)', '발견된 URL (URL)', '출처 페이지 (Source)', 'HTML 태그 (Tag)', '속성 (Attribute)'])
    df.loc[0] = ['지정 도메인 내부에서 스캔된 결과가 없거나 차단되었습니다.', '', '', '', '', '', '']
    print("⚠️ 지정 도메인에 매칭되는 데이터가 없습니다.")

with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='통합 병렬스캔 결과')
    worksheet = writer.sheets['통합 병렬스캔 결과']
    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
