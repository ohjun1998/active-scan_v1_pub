# json_to_excel.py
import json
import os
import glob
import re
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment

output_dir = 'katana_outputs'
excel_file = 'katana_multi_scan_report.xlsx'
data = []

# Secrets의 TARGET_URLS 환경변수를 읽어와 허용된 도메인 세트 생성
target_urls_env = os.environ.get('TARGET_URLS', '')
allowed_domains = set()

for line in target_urls_env.split('\n'):
    line_str = line.strip()
    if line_str:
        domain = re.sub(r'^https?://', '', line_str)
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
                    
                try:
                    if '//' in url_str:
                        target_source = url_str.split('/')[2].split(':')[0]
                    else:
                        target_source = url_str.split('/')[0].split(':')[0]
                except Exception:
                    target_source = "Unknown"
                    
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
    # 🔥 [요청 사항 1] 타겟 도메인은 내림차순(Z-A), 내부 URL은 오름차순(A-Z)으로 보기 좋게 정렬
    df = df.sort_values(by=['대상 타겟 (Target)', '발견된 URL (URL)'], ascending=[False, True])
    print(f"🟢 [필터링 완료] 총 {len(data)}개의 순정 타겟 엔드포인트가 엑셀에 매핑되었습니다!")
else:
    df = pd.DataFrame(columns=['대상 타겟 (Target)', '찾은 시간 (Timestamp)', '요청 메서드 (Method)', '발견된 URL (URL)', '출처 페이지 (Source)', 'HTML 태그 (Tag)', '속성 (Attribute)'])
    df.loc[0] = ['지정 도메인 내부에서 스캔된 결과가 없거나 차단되었습니다.', '', '', '', '', '', '']

# 최종 엑셀 파일 생성 및 초경량 가독성 서식 지정
with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='통합 병렬스캔 결과')
    worksheet = writer.sheets['통합 병렬스캔 결과']
    
    # 🔥 [요청 사항 2] 파일 크기를 키우지 않는 초경량 시각 서식 (맑은 고딕 + 연한 블루 헤더)
    header_font = Font(name='Malgun Gothic', size=11, bold=True, color='000000')
    header_fill = PatternFill(start_color='E6F0FA', end_color='E6F0FA', fill_type='solid') # 눈이 편안한 미색 블루
    data_font = Font(name='Malgun Gothic', size=10)
    center_alignment = Alignment(horizontal='center', vertical='center')
    
    # 헤더 행 서식 적용
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        
    # 데이터 행 폰트 일괄 통일 (용량에 영향 주지 않음)
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
        for cell in row:
            cell.font = data_font
            # 요청 메서드 컬럼(3번째)만 중앙 정렬하여 가독성 업
            if cell.column == 3:
                cell.alignment = center_alignment

    # 열 너비 자동 조절
    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
