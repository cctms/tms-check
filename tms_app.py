import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="TMS 수질 시험 항목 추출", layout="wide")

@st.cache_data
def load_guide_data():
    try:
        # 파일 목록에서 가이드북 찾기
        f_list = os.listdir('.')
        g_p = next((f for f in f_list if '가이드북' in f or '시험방법' in f), None)
        
        if not g_p:
            return None

        # 1. 일단 엑셀을 읽어옴
        df_raw = pd.read_excel(g_p)
        
        # 2. '일반현황'이라는 글자가 있는 행을 찾아 헤더(시험명)로 설정
        header_idx = 0
        for i, row in df_raw.iterrows():
            if "일반현황" in str(row.values):
                df_raw.columns = df_raw.iloc[i]  # 해당 행을 컬럼명(시험명)으로
                header_idx = i + 1
                break
        
        # 3. 실제 데이터 영역만 남김
        df_final = df_raw.iloc[header_idx:].reset_index(drop=True)
        
        # 4. 분류(세로 2번째 열) 병합 해제
        df_final.iloc[:, 1] = df_final.iloc[:, 1].ffill()
        
        return df_final
    except Exception as e:
        st.error(f"엑셀 로드 중 오류 발생: {e}")
        return None

df = load_guide_data()

# 'O' 표시 확인 함수 (공백 제거, 대문자 변환)
def check_mark(val):
    s = str(val).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

st.title("📋 개선내역별 수행 시험 리스트")

if df is not None:
    # 세로(개선내역) 열은 보통 3번째 열(index 2)
    search_q = st.text_input("개선내역 입력 (예: 측정기기 교체)", "")
    
    if search_q:
        # 입력한 검색어가 포함된 행(세로) 찾기
        matches = df[df.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        
        if not matches.empty:
            # 사용자가 선택할 수 있게 표시
            matches['display'] = matches.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            selected_name = st.selectbox("정확한 개선내역을 선택하세요", matches['display'].tolist())
            
            # 선택된 행(세로) 데이터 가져오기
            target_row = matches[matches['display'] == selected_name].iloc[0]
            
            st.divider()
            st.subheader(f"✅ '{selected_name}' 시 수행해야 할 시험")

            # 가로(컬럼명=시험명) 순서대로 스캔하며 'O' 표시가 있는 것만 추출
            # 순번, 분류, 개선내역 이후의 열부터 검사
            test_cols = df.columns[3:] 
            
            col1, col2, col3 = st.columns(3)
            
            # 섹션별 키워드로 구분해서 출력
            with col1:
                st.markdown("### [1. 통합시험]")
                for col in test_cols:
                    if any(k in str(col) for k in ["일반현황", "점검사항", "자료생성", "자료수집기", "관제센터"]):
                        if check_mark(target_row[col]):
                            st.info(f"📝 {col}")

            with col2:
                st.markdown("### [2. 확인검사]")
                for col in test_cols:
                    if any(k in str(col) for k in ["구조", "시료", "승인", "방법", "범위", "교정", "표준물질", "정도검사", "일자", "유량계", "반복성", "드리프트", "재현성"]):
                        if check_mark(target_row[col]):
                            st.success(f"📝 {col}")

            with col3:
                st.markdown("### [3. 상대정확도]")
                for col in test_cols:
                    if "상대" in str(col):
                        if check_mark(target_row[col]):
                            st.warning(f"📝 {col}")

else:
    st.info("폴더에 '가이드북' 또는 '시험방법' 단어가 포함된 엑셀 파일을 넣어주세요.")
