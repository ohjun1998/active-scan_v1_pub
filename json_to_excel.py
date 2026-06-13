# json_to_excel.py
import json
import os
import glob
import re
import pandas as pd

output_dir = 'katana_outputs'
excel_file = 'katana_multi_scan_report.xlsx'
data = []

# 각 VM에서 수집되어 복호화된 jsonl 파일들 매핑
jsonl_files = glob.glob(os.path.join(output_dir, 'part_*.jsonl'))

print(f"🔍 [파이썬 디버그] 탐색된 수집 파일 목록: {jsonl_files}")

if jsonl_files:
    for file_path in jsonl_files:
        print(f"📖 현재 파일 분석 중: {file_path}")
        
        # 🚀 [디버그 장치] 실제 파일 안에 뭐가 들었는지 상위 3줄만 로그에 강제 출력합니다.
        try:
            with open(file_path, 'r', encoding='utf-8') as f_test:
                sample = [f_test.readline().strip() for _ in range(3)]
                print(f"   📊 [파일 내부 샘플 데이터] 처음 3줄 내용: {sample}")
        except Exception as e:
            print(f"   ❌ 샘플 읽기 실패: {e}")

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
                
                # 1. 순정 JSON 구조 파싱 시도
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
                    pass # JSON 파싱 실패 시 아래 정규식 마스터 필터로 우회

                # 2. 🔥 [치트키] JSON 파싱이 깨졌거나 실패했어도, 문장 안에서 URL 강제 추출
                if not url_str:
                    url_match = re.search(r'(https?://[^\s"\'},]+)', line_str)
                    if url_match:
                        url_str = url_match.group(1)

                # 그래도 URL 정보가 없으면 진짜 무효한 라인이므로 스킵
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
    print(f"🟢 [성공] 총 {len(data)}개의 가용한 엔드포인트를 엑셀에 매핑했습니다!")
else:
    df = pd.DataFrame(columns=['대상 타겟 (Target)', '찾은 시간 (Timestamp)', '요청 메서드 (Method)', '발견된 URL (URL)', '출처 페이지 (Source)', 'HTML 태그 (Tag)', '속성 (Attribute)'])
    df.loc[0] = ['스캔된 결과가 없거나 차단되었습니다.', '', '', '', '', '', '']
    print("⚠️ [경고] 여전히 처리할 데이터가 없습니다. 주입된 데이터 포맷을 재점검하세요.")

# 최종 엑셀 파일 생성
with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='통합 병렬스캔 결과')
    worksheet = writer.sheets['통합 병렬스캔 결과']
    
    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
