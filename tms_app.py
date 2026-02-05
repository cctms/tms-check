import streamlit as st
import pandas as pd
from io import BytesIO
import os

st.set_page_config(page_title="TMS 수질 시험 매칭", layout="wide")

@st.cache_data
def load_all_data():
    try:
        f_list = os.listdir('.')
        g_p = next((f for f in f_list if '가이드북' in f or '시험방법' in f), None)
        r_p = next((f for f in f_list if '1.통합' in f), None)
        c_p = next((f for f in f_list if '2.확인' in f), None)
        s_p = next((f for f in f_list if '상대' in f or '3.' in f), None)
        if not g_p: return None, None, None, None
        
        # 가이드북 헤더 탐색
        guide_raw = pd.read_excel(g_p, header=None)
        header_idx = 2
        for i in range(min(10, len(guide_raw))):
            row_str = "".join(guide_raw.iloc[i].astype(str))
            if "일반현황" in row_str or "반복성" in row_str:
                header_idx = i
                break
        
        df_g = pd.read_excel(g_p, skiprows=header_idx)
        df_g.iloc[:, 1] = df_g.iloc[:, 1].ffill()
        
        r_s = pd.read_excel(r_p, sheet_name=None) if r_p else {}
        c_s = pd.read_excel(c_p, sheet_name=None) if c_p else {}
        s_s = pd.read_excel(s_p, sheet_name=None) if s_p else {}
        
        return df_g, r_s, c_s, s_s
    except:
        return None, None, None, None

df_g, r_s, c_s, s_s = load_all_data()

def is_ok(v):
    if pd.isna(v): return False
    s = str(v).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

st.title("📋 수질 TMS 수행항목 및 상세 데이터")

if df_g is not None:
    search_q = st.text_input("개선내역 입력 (예: 측정기기 교체)", "")
    if search_q:
        match_rows = df_g[df_g.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        if not match_rows.empty:
            match_rows['dn'] = match_rows.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            sel = st.selectbox("항목 선택", ["선택하세요"] + match_rows['dn'].tolist())
            
            if sel != "선택하세요":
                row = match_rows[match_rows['dn'] == sel].iloc[0]

                # 파일별 키워드 매칭 로직
                def show_data(test_name, sheets, f_type):
                    tn = test_name.replace(" ", "")
                    for sn in sheets.keys():
                        sn_c = str(sn).replace(" ", "")
                        # 매칭 조건 (이름 포함 혹은 특수 규칙)
                        match = (tn in sn_c or sn_c in tn)
                        if not match and f_type == "통합" and any(k in tn for k in ["점검", "생성", "수집기"]):
                            match = any(k in sn_c for k in ["점검", "생성", "전송", "관제"])
                        if not match and f_type == "확인" and any(k in tn for k in ["구조", "외관"]):
                            match = any(k in sn_c for k in ["구조", "시료", "승인", "방법", "범위", "교정", "일자"])
                        
                        if match:
                            st.dataframe(sheets[sn].fillna(""), use_container_width=True)

                r_keys = ["일반현황", "점검사항", "자료생성", "자료수집기", "관제센터"]
                c_keys = ["구조", "시료", "승인", "방법", "범위", "교정", "표준물질", "정도검사", "교정일자", "유량계", "누적값", "반복성", "드리프트", "재현성"]

                col1, col2, col3 = st.columns(3)

                with col1:
                    st.header("1. 통합시험")
                    for col in df_g.columns:
                        if any(k in str(col) for k in r_keys) and is_ok(row[col]):
                            st.markdown(f"### 📍 {col}")
                            show_data(str(col), r_s, "통합")

                with col2:
                    st.header("2. 확인검사")
                    for col in df_g.columns:
                        if any(k in str(col) for k in c_keys) and is_ok(row[col]):
                            st.markdown(f"### 📍 {col}")
                            show_data(str(col), c_s, "확인")

                with col3:
                    st.header("3. 상대정확도")
                    for col in df_g.columns:
                        if "상대" in str(col) and is_ok(row[col]):
                            st.markdown(f"### 📍 {col}")
                            for sn in s_s.keys():
                                st.dataframe(s_s[sn].fillna(""), use_container_width=True)
