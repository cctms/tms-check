import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="TMS 수질 시험 매칭", layout="wide")

@st.cache_data
def load_data():
    try:
        f_list = os.listdir('.')
        g_p = next((f for f in f_list if '가이드북' in f or '시험방법' in f), None)
        r_p = next((f for f in f_list if '1.통합' in f), None)
        c_p = next((f for f in f_list if '2.확인' in f), None)
        s_p = next((f for f in f_list if '상대' in f or '3.' in f), None)
        
        # 가이드북 로드 (항목명이 있는 행을 정확히 타겟팅)
        guide_raw = pd.read_excel(g_p, header=None)
        header_row = 0
        for i in range(len(guide_raw)):
            if "일반현황" in str(guide_raw.iloc[i].values):
                header_row = i
                break
        
        df_g = pd.read_excel(g_p, skiprows=header_row)
        df_g.iloc[:, 1] = df_g.iloc[:, 1].ffill()
        
        # 각 파일별 시트 데이터 로드
        r_s = pd.read_excel(r_p, sheet_name=None) if r_p else {}
        c_s = pd.read_excel(c_p, sheet_name=None) if c_p else {}
        s_s = pd.read_excel(s_p, sheet_name=None) if s_p else {}
        
        return df_g, r_s, c_s, s_s
    except:
        return None, None, None, None

df_g, r_s, c_s, s_s = load_data()

def is_ok(v):
    s = str(v).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

st.title("📋 수질 TMS 수행항목 매칭")

if df_g is not None:
    search_q = st.text_input("개선내역 입력", "")
    if search_q:
        match_rows = df_g[df_g.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        if not match_rows.empty:
            match_rows['dn'] = match_rows.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            sel = st.selectbox("개선내역 선택", ["선택하세요"] + match_rows['dn'].tolist())
            
            if sel != "선택하세요":
                row = match_rows[match_rows['dn'] == sel].iloc[0]
                
                # 가이드북 열 순서대로 분류 기준 설정
                # 실제 엑셀 열 이름들을 리스트로 추출 (순서 보장)
                all_cols = [c for c in df_g.columns if not any(ex in str(c) for ex in ["순번", "분류", "개선내역", "Unnamed"])]

                col1, col2, col3 = st.columns(3)

                with col1:
                    st.header("1. 통합시험")
                    for c_name in all_cols:
                        # 가이드북 열 이름에 통합시험 관련 키워드가 있고 'ㅇ' 체크된 경우
                        if any(k in str(c_name) for k in ["일반현황", "점검사항", "자료생성", "자료수집기", "관제센터"]):
                            if is_ok(row[c_name]):
                                st.subheader(f"📍 {c_name}") # 가이드북 항목명 출력
                                # 해당 항목명과 유사한 탭 검색하여 데이터 출력
                                for sn in r_s.keys():
                                    if str(sn).replace(" ","") in str(c_name).replace(" ","") or str(c_name).replace(" ","") in str(sn).replace(" ",""):
                                        st.dataframe(r_s[sn].fillna(""), use_container_width=True)

                with col2:
                    st.header("2. 확인검사")
                    for c_name in all_cols:
                        if any(k in str(c_name) for k in ["구조", "시료", "승인", "방법", "범위", "교정", "표준물질", "정도검사", "교정일자", "유량계", "누적값", "반복성", "드리프트", "재현성"]):
                            if is_ok(row[c_name]):
                                st.subheader(f"📍 {c_name}")
                                for sn in c_s.keys():
                                    if str(sn).replace(" ","") in str(c_name).replace(" ","") or str(c_name).replace(" ","") in str(sn).replace(" ",""):
                                        st.dataframe(c_s[sn].fillna(""), use_container_width=True)

                with col3:
                    st.header("3. 상대정확도")
                    for c_name in all_cols:
                        if "상대" in str(c_name) and is_ok(row[c_name]):
                            st.subheader(f"📍 {c_name}")
                            for sn in s_s.keys():
                                st.dataframe(s_s[sn].fillna(""), use_container_width=True)
