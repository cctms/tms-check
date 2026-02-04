import pandas as pd
import io

def generate_inspection_excel(user_input):
    # 1. 파일 로드 (제공된 CSV 파일 기준)
    try:
        df_guide = pd.read_csv('개선내역에 따른 시험방법(2025 최종).xlsx - ★최종(가이드북).csv', skiprows=1)
    except FileNotFoundError:
        return "가이드북 파일을 찾을 수 없습니다."

    # 2. 개선 내역 검색
    match = df_guide[df_guide['개선 내역'].str.contains(user_input, na=False, case=False)]
    if match.empty:
        return f"'{user_input}'에 해당하는 개선 내역을 찾을 수 없습니다."

    # 3. 필요한 시험 항목 번호 추출 (1.일반현황 ~ 8.관제센터)
    test_columns = [
        '1. 일반현황', '2. 하드웨어 규격', '3. 소프트웨어 \n기능 규격', 
        '4. 자료정의', '5. 측정기기 \n점검사항', '6. 자료생성', 
        '7. 측정기기-자료수집기', '8. 자료수집기-관제센터'
    ]
    
    selected_row = match.iloc[0]
    required_tests = [col.split('.')[0].strip() for col in test_columns if str(selected_row[col]).startswith('O')]

    # 4. 각 시험별 상세 조사표 데이터 매핑 (조사표 파일 분석 내용 반영)
    # 실제 환경에서는 각 번호별 CSV/Excel 파일을 로드하도록 구현합니다.
    test_details = {
        "1": "1.1 장비 설치현황(S/N 확인), 1.2 VPN 장비 정보 일치 여부",
        "2": "2.1 직렬포트 할당(기기당 1개), 2.2 자료보안 안정성(외부변조 방지)",
        "3": "3.1 수집기능, 3.2 저장기능(30일 이상), 3.5 비밀번호 설정 기능",
        "4": "4.1 5분자료 생성 방식, 4.2 시간자료 산출 적정성",
        "5": "5.1 측정상수(Factor/Offset) 전송, 5.3 로그기록 저장, 5.4 비밀번호 설정",
        "6": "6.1 상태정보 코드 구성, 6.2 상태정보 우선순위 준수 여부",
        "7": "7.1 실시간 자료전송 전문 형식, 7.3 오류처리(3회 재시도 및 시간초과)",
        "8": "8.1 인증된 IP 접속 응답, 8.2 전송범위 제한(2시간), 8.7 시간변경 처리"
    }

    # 5. 결과 데이터프레임 구성
    result_data = []
    for test_no in required_tests:
        result_data.append({
            "개선내역": selected_row['개선 내역'],
            "시험구분": f"{test_no}번 시험",
            "상세 점검 항목": test_details.get(test_no, "상세 가이드북 참조"),
            "비고": selected_row.get('참 고', '-')
        })

    df_result = pd.DataFrame(result_data)

    # 6. 엑셀 파일로 내보내기
    output_filename = f"조사표_{user_input.replace(' ', '_')}.xlsx"
    df_result.to_excel(output_filename, index=False)
    
    print(f"✅ '{output_filename}' 파일이 성공적으로 생성되었습니다.")
    return df_result

# 사용 예시
user_query = input("개선 내용을 입력하세요 (예: 측정기기 교체): ")
result = generate_inspection_excel(user_query)
if isinstance(result, pd.DataFrame):
    print("\n[생성된 조사표 요약]")
    print(result[['시험구분', '상세 점검 항목']])