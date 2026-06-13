import json
import os
import glob
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

output_dir = 'katana_outputs'
excel_file = 'katana_multi_scan_report.xlsx'
data = []

# 1. Secrets의 TARGET_URLS 환경변수를 읽어와 허용된 도메인 세트 생성
target_urls_env = os.environ.get('TARGET_URLS', '')
allowed_domains = set()

for line in target_urls_env.split('\n'):
    line_str = line.strip()
    if line_str:
        domain = re.sub(r'^https?://', '', line_str)
        domain = domain.split('/')[0].split(':')[0]
        allowed_domains.add(domain)

print(f"🎯 [필터 켜짐] 다음 지정 도메인 결과만 엑셀에 저장됩니다: {allowed_domains}")

# 2. 각 VM에서 수집되어 복호화된 jsonl 파일들 매핑 및 파싱
jsonl_files = glob.glob(os.path.join(output_dir, 'part_*.jsonl'))

if jsonl_files:
    for file_path in jsonl_files:
        # errors='ignore' 주입으로 유니코드 깨짐 크래시 원천 차단
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
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
                        source = obj.get('source', '')  # Katana 원본의 소스(출처) 데이터 매핑
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
                    '출처 페이지 (Source)': source,  # 데이터프레임 컬럼 정상 활성화
                    'HTML 태그 (Tag)': tag,
                    '속성 (Attribute)': attribute
                }
                data.append(row)

# 3. 데이터프레임 빌드 및 내림차순/오름차순 정렬
if data:
    df = pd.DataFrame(data)
    df = df.sort_values(by=['대상 타겟 (Target)', '발견된 URL (URL)'], ascending=[False, True])
    print(f"🟢 [필터링 완료] 총 {len(data)}개의 순정 타겟 엔드포인트가 엑셀에 매핑되었습니다!")
else:
    df = pd.DataFrame(columns=['대상 타겟 (Target)', '찾은 시간 (Timestamp)', '요청 메서드 (Method)', '발견된 URL (URL)', '출처 페이지 (Source)', 'HTML 태그 (Tag)', '속성 (Attribute)'])
    df.loc[0] = ['지정 도메인 내부에서 스캔된 결과가 없거나 차단되었습니다.', '', '', '', '', '', '']

# --- 📊 엑셀 엔진 레이아웃 빌드 조립 단 ---
wb = Workbook()

font_family = 'Malgun Gothic'
header_font = Font(name=font_family, size=11, bold=True, color='000000')
header_fill = PatternFill(start_color='E6F0FA', end_color='E6F0FA', fill_type='solid')
total_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
data_font = Font(name=font_family, size=10)
center_alignment = Alignment(horizontal='center', vertical='center')
left_alignment = Alignment(horizontal='left', vertical='center')

# A. 대시보드 메인 현황판 수립
ws_dashboard = wb.active
ws_dashboard.title = "대시보드"
ws_dashboard.append(["대상 타겟 도메인 (Target Domain)", "수집된 고유 URL 개수", "상세 페이지 바로가기"])

for cell in ws_dashboard[1]:
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_alignment
ws_dashboard.row_dimensions[1].height = 24

# B. 도메인 분할 이식 및 하이퍼링크 크로스를 통한 중앙정렬 연산
total_url_count = 0

if data:
    unique_domains = df['대상 타겟 (Target)'].unique()
    last_row_idx = 1
    
    for idx, domain in enumerate(unique_domains, start=2):
        sheet_title = domain[:30]
        ws_domain = wb.create_sheet(title=sheet_title)
        
        # 1) 초경량 상단 대시보드 복귀 단추 배너 주입 (A1~H1 병합으로 스펙 업)
        ws_domain.merge_cells("A1:H1")
        back_btn = ws_domain["A1"]
        back_btn.value = "⬅️ 대시보드 현황판으로 돌아가기"
        back_btn.hyperlink = "#'대시보드'!A1"
        back_btn.font = Font(name=font_family, size=11, bold=True, color='FFFFFF')
        back_btn.fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
        back_btn.alignment = center_alignment
        ws_domain.row_dimensions[1].height = 26
        
        # 2) 소스 데이터 이식
        domain_subset = df[df['대상 타겟 (Target)'] == domain]
        headers = list(df.columns)
        ws_domain.append(headers)
        
        for r in dataframe_to_rows(domain_subset, index=False, header=False):
            ws_domain.append(r)
            
        # 데이터 탭 테이블 헤더 서식 지정 (2행)
        for cell in ws_domain[2]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_alignment
        ws_domain.row_dimensions[2].height = 20
            
        # 데이터 바디 서식 정렬 일괄 전개
        for row in ws_domain.iter_rows(min_row=3, max_row=ws_domain.max_row):
            for cell in row:
                cell.font = data_font
                # 가독성 설계: 타겟, 시간, 메서드, 태그, 속성은 가운데 정렬 / 발견 URL과 출처(Source)는 왼쪽 정렬
                if cell.column in [1, 2, 3, 6, 7]:
                    cell.alignment = center_alignment
                else:
                    cell.alignment = left_alignment
                    
                # 긴 긴 URL 텍스트 가독성을 위해 셀 내부 줄바꿈 허용 세팅 보강
                if cell.column in [4, 5]:
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
                    
        # 열 너비 수동 연산 (병합 크래시 우회 엔진 고정)
        for col_idx in range(1, len(headers) + 1):
            max_len = 0
            for row_idx in range(2, ws_domain.max_row + 1):
                val = str(ws_domain.cell(row=row_idx, column=col_idx).value or '')
                if len(val) > max_len:
                    max_len = len(val)
            col_letter = get_column_letter(col_idx)
            # URL과 출처 페이지 열은 데이터가 길기 때문에 너비를 넉넉하게 보정
            if col_idx in [4, 5]:
                ws_domain.column_dimensions[col_letter].width = max(max_len + 4, 35)
            else:
                ws_domain.column_dimensions[col_letter].width = max(max_len + 3, 14)
            
        # 3) 대시보드 데이터 채우기 및 완벽 정중앙 정렬
        current_domain_count = len(domain_subset)
        total_url_count += current_domain_count
        
        ws_dashboard.cell(row=idx, column=1, value=domain).alignment = center_alignment
        ws_dashboard.cell(row=idx, column=2, value=current_domain_count).alignment = center_alignment
        
        link_cell = ws_dashboard.cell(row=idx, column=3, value="🔍 상세 내역 시트로 탭 이동")
        link_cell.hyperlink = f"#'{sheet_title}'!A2"
        link_cell.font = Font(name=font_family, size=10, color='0056B3', underline='single')
        link_cell.alignment = center_alignment
        
        ws_dashboard.row_dimensions[idx].height = 22
        ws_dashboard.cell(row=idx, column=1).font = data_font
        ws_dashboard.cell(row=idx, column=2).font = Font(name=font_family, size=10, bold=True)
        last_row_idx = idx

    # 대시보드 최하단에 전체 URL 총합 요약 행 주입
    total_row_idx = last_row_idx + 1
    ws_dashboard.cell(row=total_row_idx, column=1, value="📊 수집된 URL 총합 (Total)").alignment = center_alignment
    ws_dashboard.cell(row=total_row_idx, column=1).font = Font(name=font_family, size=10, bold=True, color='FF0000')
    ws_dashboard.cell(row=total_row_idx, column=1).fill = total_fill
    
    ws_dashboard.cell(row=total_row_idx, column=2, value=total_url_count).alignment = center_alignment
    ws_dashboard.cell(row=total_row_idx, column=2).font = Font(name=font_family, size=11, bold=True, color='FF0000')
    ws_dashboard.cell(row=total_row_idx, column=2).fill = total_fill
    
    ws_dashboard.cell(row=total_row_idx, column=3, value="").fill = total_fill
    ws_dashboard.row_dimensions[total_row_idx].height = 24

else:
    ws_empty = wb.create_sheet(title="결과 없음")
    ws_empty.append(list(df.columns))
    ws_empty.append(['지정 도메인 내부에서 스캔된 결과가 없거나 차단되었습니다.', '', '', '', '', '', ''])
    
    ws_dashboard.cell(row=2, column=1, value="N/A").alignment = center_alignment
    ws_dashboard.cell(row=2, column=2, value=0).alignment = center_alignment
    ws_dashboard.cell(row=2, column=3, value="데이터 없음").alignment = center_alignment

ws_dashboard.column_dimensions['A'].width = 38
ws_dashboard.column_dimensions['B'].width = 24
ws_dashboard.column_dimensions['C'].width = 28

wb.save(excel_file)
print(f"🏁 [출처 표기 완료] 대시보드 총합 및 소스 맵 연동 완료: {excel_file}")
