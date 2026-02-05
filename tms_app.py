import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="TMS 수질 시험 매칭", layout="wide")

@st.cache_data
def load_all():
    try:
        f_list = os.listdir('.')
        g_p = next((f for f in f_list if '가이드북' in f or '시험방법' in f), None)
        r_p = next((f for f in f_list if '1.통합' in f), None)
        c_p = next((f for f in f_list if '2.확인' in f), None)
        s_p = next((f for f in f_list if '상대' in f or '3.' in f), None)
        
        if not g_p: return None, None, None, None
        
        # 가이드북 로드 (헤더를 찾지 못할 경우를 대비해 0행부터 읽은 뒤 처리)
        df_g = pd.read_excel(g_p)
        # 실제 데이터가 시작되는 행 찾기 (일반현황이라는 글자가 있는 행)
        start_idx = 0
        for i, row in df_g.iterrows():
            if "일반현황" in str(row.values):
                # 해당 행을 컬럼명으로 재설정
                df_g.columns = df_g.iloc[i]
                start_idx = i + 1
                break
        df_g = df_g.iloc[start_idx:].reset_index(drop=True)
        df_g.iloc[:, 1] = df_g.iloc[:, 1].ffill() # 대분류 채우기
        
        r_s = pd.read_excel(r_p, sheet_name=None) if r_p else {}
        c_s = pd.read_excel(c_p, sheet_name=None) if c_p else {}
        s_s = pd.read_excel(s_p, sheet_name=None) if s_p else {}
        
        return df_g, r_s, c_s, s_s
    except Exception as e:
        st.error(f"파일 로드 오류: {e}")
        return None, None, None, None

df_g, r_s, c_s, s_s = load_all()

def is_target(val):
    s = str(val).upper().replace(" ", "")
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

st.title("📋 수질 TMS 수행항목 리스트 (가이드북 순서)")

if df_g is not None:
    # 3번 열(개선내역) 기준으로 검색
    search_q = st.text_input("개선내역 입력 (예: 측정기기 교체)", "")
    
    if search_q:
        # 검색 필터링
        matches = df_g[df_g.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        
        if not matches.empty:
            matches['display_name'] = matches.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            selected = st.selectbox("정확한 항목 선택", ["선택하세요"] + matches['display_name'].tolist())
            
            if selected != "선택하세요":
                row_data = matches[matches['display_name'] == selected].iloc[0]
                
                # 가이드북 열들을 순회하며 체크된 항목 출력
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.header("1. 통합시험")
                    # 통합시험 키워드 (가이드북 열 이름에 이 글자가 포함되면)
                    r_keys = ["일반현황", "점검사항", "자료생성", "자료수집기", "관제센터"]
                    for col_name in df_g.columns:
                        c_str = str(col_name)
                        if any(k in c_str for k in r_keys):
                            if is_target(row_data[col_name]):
                                st.subheader(f"📍 {c_str}")
                                # 데이터 매칭 출력 (탭 이름과 비교)
                                for s_name in r_s.keys():
                                    if s_name.replace(" ","") in c_str.replace(" ","") or c_str.replace(" ","") in s_name.replace(" ",""):
                                        st.dataframe(r_s[s_name].fillna(""), use_container_width=True)

                with col2:
                    st.header("2. 확인검사")
                    c_keys = ["구조", "시료", "승인", "방법", "범위", "교정", "표준물질", "정도검사", "일자", "유량계", "누적값", "반복성", "드리프트", "재현성"]
                    for col_name in df_g.columns:
                        c_str = str(col_name)
                        if any(k in c_str for k in c_keys):
                            if is_target(row_data[col_name]):
                                st.subheader(f"📍 {c_str}")
                                for s_name in c_s.keys():
                                    if s_name.replace(" ","") in c_str.replace(" ","") or c_str.replace(" ","") in s_name.replace(" ",""):
                                        st.dataframe(c_s[s_name].fillna(""), use_container_width=True)

                with col3:
                    st.header("3. 상대정확도")
                    for col_name in df_g.columns:
                        c_str = str(col_name)
                        if "상대" in c_str:
                            if is_target(row_data[col_name]):
                                st.subheader(f"📍 {c_str}")
                                for s_name in s_s.keys():
                                    st.dataframe(s_s[s_name].fillna(""), use_container_width=True)
else:
    st.warning("파일을 읽어오지 못했습니다. 파일명에 '가이드북' 또는 '시험방법'이 포함되어 있는지 확인해주세요.")
