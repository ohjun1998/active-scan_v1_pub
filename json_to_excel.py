import json
import os
import glob
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

output_dir = 'katana_outputs'
excel_file = 'katana_multi_scan_report.xlsx'
data = []

# 1. Secrets의 TARGET_URLS 환경변수를 읽어와 허용된 도메인 세트 생성 (기존 유지)
target_urls_env = os.environ.get('TARGET_URLS', '')
allowed_domains = set()

for line in target_urls_env.split('\n'):
    line_str = line.strip()
    if line_str:
        domain = re.sub(r'^https?://', '', line_str)
        domain = domain.split('/')[0].split(':')[0]
        allowed_domains.add(domain)

print(f"🎯 [필터 켜짐] 다음 지정 도메인 결과만 엑셀에 저장됩니다: {allowed_domains}")

# 2. 각 VM에서 수집되어 복호화된 jsonl 파일들 매핑 및 파싱 (기존 유지)
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

# 3. 데이터프레임 빌드 및 내림차순/오름차순 정렬 (기존 유지)
if data:
    df = pd.DataFrame(data)
    df = df.sort_values(by=['대상 타겟 (Target)', '발견된 URL (URL)'], ascending=[False, True])
    print(f"🟢 [필터링 완료] 총 {len(data)}개의 순정 타겟 엔드포인트가 엑셀에 매핑되었습니다!")
else:
    df = pd.DataFrame(columns=['대상 타겟 (Target)', '찾은 시간 (Timestamp)', '요청 메서드 (Method)', '발견된 URL (URL)', '출처 페이지 (Source)', 'HTML 태그 (Tag)', '속성 (Attribute)'])
    df.loc[0] = ['지정 도메인 내부에서 스캔된 결과가 없거나 차단되었습니다.', '', '', '', '', '', '']

# --- 🔥 [새롭게 추가/수정된 대시보드 및 멀티탭 서식 설계 단] ---

# openpyxl 워크북 수동 빌드 구동
wb = Workbook()

# 서식 옵션 정의 (기존 유저님 스타일 스펙 그대로 이식)
header_font = Font(name='Malgun Gothic', size=11, bold=True, color='000000')
header_fill = PatternFill(start_color='E6F0FA', end_color='E6F0FA', fill_type='solid') # 눈이 편안한 미색 블루
data_font = Font(name='Malgun Gothic', size=10)
link_font = Font(name='Malgun Gothic', size=10, color='0000FF', underline='single') # 대시보드 하이퍼링크용 파란색 밑줄
center_alignment = Alignment(horizontal='center', vertical='center')

# A. 대시보드 탭 생성 및 기초 공사
ws_dashboard = wb.active
ws_dashboard.title = "대시보드"

ws_dashboard.append(["대상 타겟 도메인 (Target Domain)", "수집된 고유 URL 개수", "상세 페이지 바로가기"])

# 대시보드 헤더 스타일링
for cell in ws_dashboard[1]:
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_alignment

# B. 도메인 추출 후 탭 분할 및 하이퍼링크 크로스 매핑 연산
if data:
    unique_domains = df['대상 타겟 (Target)'].unique()
    
    for idx, domain in enumerate(unique_domains, start=2):
        # 1) 해당 도메인 글자로 새 시트(탭) 동적 바인딩
        # 단, 엑셀 시트명 제약 조건(최대 31자 및 특수문자 제한) 방어를 위해 slice 처리
        sheet_title = domain[:30]
        ws_domain = wb.create_sheet(title=sheet_title)
        
        # 2) 해당 도메인 알짜배기 소스 데이터만 추출 (Subset)
        domain_subset = df[df['대상 타겟 (Target)'] == domain]
        
        # 3) 컬럼 헤더 주입 및 데이터 로우 단위 안전 이식
        headers = list(df.columns)
        ws_domain.append(headers)
        
        for r in dataframe_to_rows(domain_subset, index=False, header=False):
            ws_domain.append(r)
            
        # 4) 개별 도메인 시트 내부 가독성 서식 일괄 적용 (기존 스타일 이식 루프)
        for cell in ws_domain[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_alignment
            
        for row in ws_domain.iter_rows(min_row=2, max_row=ws_domain.max_row):
            for cell in row:
                cell.font = data_font
                if cell.column == 3: # Method 컬럼 중앙 정렬
                    cell.alignment = center_alignment
                    
        # 5) 개별 도메인 시트 열 너비 자동 조절 (기존 로직 보완 이식)
        for col in ws_domain.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            ws_domain.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
        # 6) 🔥 [핵심 요청 기능] 대시보드 데이터 생성 및 동적 하이퍼링크 주입
        ws_dashboard.cell(row=idx, column=1, value=domain).font = data_font
        ws_dashboard.cell(row=idx, column=2, value=len(domain_subset)).font = data_font
        ws_dashboard.cell(row=idx, column=2).alignment = center_alignment
        
        # 하이퍼링크 객체 조립: 클릭하면 해당 시트의 A1 셀로 화면 점프 연동
        link_cell = ws_dashboard.cell(row=idx, column=3, value="👉 클릭하여 해당 탭으로 이동")
        link_cell.hyperlink = f"#'{sheet_title}'!A1"
        link_cell.font = link_font
        link_cell.alignment = center_alignment

else:
    # 데이터가 아예 없을 때의 방어 예외 시트 빌드 (기존 유지 반영)
    ws_empty = wb.create_sheet(title="결과 없음")
    headers = ['대상 타겟 (Target)', '찾은 시간 (Timestamp)', '요청 메서드 (Method)', '발견된 URL (URL)', '출처 페이지 (Source)', 'HTML 태그 (Tag)', '속성 (Attribute)']
    ws_empty.append(headers)
    ws_empty.append(['지정 도메인 내부에서 스캔된 결과가 없거나 차단되었습니다.', '', '', '', '', '', ''])
    
    ws_dashboard.cell(row=2, column=1, value="N/A").font = data_font
    ws_dashboard.cell(row=2, column=2, value=0).font = data_font
    ws_dashboard.cell(row=2, column=3, value="데이터 없음").font = data_font

# 대시보드 탭 열 너비 이쁘게 보정
for col in ws_dashboard.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    ws_dashboard.column_dimensions[col[0].column_letter].width = max(max_len + 5, 25)

# 4. 파일 최종 세이브
wb.save(excel_file)
print(f"🏁 [최종 성공] 대시보드 하이퍼링크 연동 엑셀 리포트 패키징 완료: {excel_file}")
