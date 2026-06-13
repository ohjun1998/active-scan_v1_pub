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

# --- 📊 [초경량 고가독성 엑셀 리모델링 엔진] ---
wb = Workbook()

# 글로벌 공통 서식 옵션 정의
font_family = 'Malgun Gothic'
header_font = Font(name=font_family, size=11, bold=True, color='000000')
header_fill = PatternFill(start_color='E6F0FA', end_color='E6F0FA', fill_type='solid') # 눈이 편안한 미색 블루
data_font = Font(name=font_family, size=10)
center_alignment = Alignment(horizontal='center', vertical='center')
left_alignment = Alignment(horizontal='left', vertical='center')

# A. 대시보드 탭 생성 및 정중앙 정렬 세팅
ws_dashboard = wb.active
ws_dashboard.title = "대시보드"

ws_dashboard.append(["대상 타겟 도메인 (Target Domain)", "수집된 고유 URL 개수", "상세 페이지 바로가기"])

# 대시보드 헤더 서식 지정
for cell in ws_dashboard[1]:
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_alignment

# B. 도메인별 탭 분할 및 상호 링크 시스템 연산
if data:
    unique_domains = df['대상 타겟 (Target)'].unique()
    
    for idx, domain in enumerate(unique_domains, start=2):
        sheet_title = domain[:30] # 엑셀 글자수 제약 31자 방어
        ws_domain = wb.create_sheet(title=sheet_title)
        
        # 1) 🔥 [요청 사항] 각 도메인 탭 최상단 A1 셀에 대시보드 복귀 초경량 링크 버튼 주입
        ws_domain.merge_cells("A1:G1")
        back_btn = ws_domain["A1"]
        back_btn.value = "⬅️ 대시보드 현황판으로 돌아가기"
        back_btn.hyperlink = "#'대시보드'!A1"
        back_btn.font = Font(name=font_family, size=11, bold=True, color='FFFFFF')
        back_btn.fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid') # 깔끔한 블루 버튼 색상
        back_btn.alignment = center_alignment
        ws_domain.row_dimensions[1].height = 26 # 버튼 행 높이 확보
        
        # 데이터 서브셋 추출 및 주입 (2번째 행부터 헤더 시작)
        domain_subset = df[df['대상 타겟 (Target)'] == domain]
        headers = list(df.columns)
        ws_domain.append(headers) # 2번째 행에 자동 적재됨
        
        for r in dataframe_to_rows(domain_subset, index=False, header=False):
            ws_domain.append(r)
            
        # 2) 도메인 시트 헤더(2번째 행) 및 데이터 정렬/서식 적용
        for cell in ws_domain[2]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_alignment
        ws_domain.row_dimensions[2].height = 20
            
        for row in ws_domain.iter_rows(min_row=3, max_row=ws_domain.max_row):
            for cell in row:
                cell.font = data_font
                if cell.column in [1, 2, 3, 6, 7]: # 도메인, 타임스탬프, 메서드, 태그 등은 중앙 정렬
                    cell.alignment = center_alignment
                else:
                    cell.alignment = left_alignment
                    
        # 열 너비 자동 조절 (복귀 버튼이 있는 1행은 계산에서 제외)
        for col in ws_domain.columns:
            max_len = max(len(str(cell.value or '')) for cell in col[1:]) # 2번째 행부터 너비 계산
            col_letter = col[0].column_letter
            ws_domain.column_dimensions[col_letter].width = max(max_len + 3, 14)
            
        # 3) 🔥 [요청 사항] 대시보드 데이터 삽입 및 글자 정중앙 정렬
        ws_dashboard.cell(row=idx, column=1, value=domain).alignment = center_alignment
        ws_dashboard.cell(row=idx, column=2, value=len(domain_subset)).alignment = center_alignment
        
        link_cell = ws_dashboard.cell(row=idx, column=3, value="🔍 상세 내역 시트로 탭 이동")
        link_cell.hyperlink = f"#'{sheet_title}'!A2" # 탭 이동 시 2행(헤더 위치)으로 스크롤 이동
        link_cell.font = Font(name=font_family, size=10, color='0056B3', underline='single')
        link_cell.alignment = center_alignment
        
        ws_dashboard.row_dimensions[idx].height = 22
        
        # 대시보드 폰트 적용
        ws_dashboard.cell(row=idx, column=1).font = data_font
        ws_dashboard.cell(row=idx, column=2).font = Font(name=font_family, size=10, bold=True)

else:
    ws_empty = wb.create_sheet(title="결과 없음")
    ws_empty.append(list(df.columns))
    ws_empty.append(['지정 도메인 내부에서 스캔된 결과가 없거나 차단되었습니다.', '', '', '', '', '', ''])
    
    ws_dashboard.cell(row=2, column=1, value="N/A").alignment = center_alignment
    ws_dashboard.cell(row=2, column=2, value=0).alignment = center_alignment
    ws_dashboard.cell(row=2, column=3, value="데이터 없음").alignment = center_alignment

# 대시보드 자체 열 너비 가독성 최적화
ws_dashboard.column_dimensions['A'].width = 38
ws_dashboard.column_dimensions['B'].width = 24
ws_dashboard.column_dimensions['C'].width = 28

# 4. 파일 최종 세이브
wb.save(excel_file)
print(f"🏁 [리모델링 완료] 대시보드 중앙 정렬 및 복귀 버튼 설계 완수: {excel_file}")
