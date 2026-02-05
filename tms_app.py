import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="TMS 수질 시험 매칭", layout="wide")

@st.cache_data
def load_all_data():
    try:
        f_list = os.listdir('.')
        g_p = next((f for f in f_list if '가이드북' in f or '시험방법' in f), None)
        if not g_p: return None
        
        # 가이드북 로드
        guide_raw = pd.read_excel(g_p, header=None)
        header_idx = 0
        for i in range(min(10, len(guide_raw))):
            row_str = "".join(guide_raw.iloc[i].astype(str))
            if "일반현황" in row_str:
                header_idx = i
                break
        
        df_g = pd.read_excel(g_p, skiprows=header_idx)
        df_g.iloc[:, 1] = df_g.iloc[:, 1].ffill() # 대분류 채우기
        return df_g
    except:
        return None

df_g = load_all_data()

def is_checked(v):
    if pd.isna(v): return False
    s = str(v).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

st.title("📋 수질 TMS 수행항목 리스트")

if df_g is not None:
    search_q = st.text_input("개선내역 입력 (예: 측정기기 교체)", "")
    
    if search_q:
        match_rows = df_g[df_g.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        
        if not match_rows.empty:
            match_rows['dn'] = match_rows.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            sel = st.selectbox("항목 선택", ["선택하세요"] + match_rows['dn'].tolist())
            
            if sel != "선택하세요":
                target_row = match_rows[match_rows['dn'] == sel].iloc[0]
                
                # 1. 통합시험 리스트 (가이드북 순서대로)
                r_cols = ["일반현황", "점검사항", "자료생성", "자료수집기", "관제센터"]
                # 2. 확인검사 리스트 (가이드북 순서대로)
                c_cols = ["시료채취조", "형식승인", "측정방법", "측정범위", "교정기능", "표준물질", "정도검사", "교정일자", "유량계", "누적값", "반복성", "드리프트", "재현성"]
                # 3. 상대정확도
                s_cols = ["상대정확도"]

                c1, c2, c3 = st.columns(3)

                with c1:
                    st.header("1. 통합시험")
                    for col in df_g.columns:
                        if any(k in str(col) for k in r_cols):
                            if is_checked(target_row[col]):
                                st.write(f"✅ {col}")

                with c2:
                    st.header("2. 확인검사")
                    for col in df_g.columns:
                        if any(k in str(col) for k in c_cols):
                            if is_checked(target_row[col]):
                                st.write(f"✅ {col}")

                with c3:
                    st.header("3. 상대정확도")
                    for col in df_g.columns:
                        if any(k in str(col) for k in s_cols):
                            if is_checked(target_row[col]):
                                st.write(f"✅ {col}")
else:
    st.error("가이드북 엑셀 파일을 찾을 수 없습니다.")
