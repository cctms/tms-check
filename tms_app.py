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
        
        # 가이드북 로드 (항목명이 있는 3행을 헤더로 설정)
        guide_raw = pd.read_excel(g_p, header=None)
        header_idx = 2
        for i in range(min(6, len(guide_raw))):
            row_str = str(guide_raw.iloc[i].values)
            if "일반현황" in row_str or "반복성" in row_str:
                header_idx = i
                break
        
        df_g = pd.read_excel(g_p, skiprows=header_idx)
        df_g.iloc[:, 1] = df_g.iloc[:, 1].ffill() # 대분류 병합 해제
        
        r_s = pd.read_excel(r_p, sheet_name=None) if r_p else {}
        c_s = pd.read_excel(c_p, sheet_name=None) if c_p else {}
        s_s = pd.read_excel(s_p, sheet_name=None) if s_p else {}
        
        return df_g, r_s, c_s, s_s
    except Exception as e:
        return None, None, None, None

df_g, r_s, c_s, s_s = load_all_data()

def is_checked(v):
    if pd.isna(v): return False
    s = str(v).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

st.title("📋 수질 TMS 수행항목 매칭 (엑셀 순서 기준)")

if df_g is not None:
    search_q = st.text_input("개선내역 입력 (예: 측정기기 교체)", "")
    
    if search_q:
        match_rows = df_g[df_g.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        
        if not match_rows.empty:
            match_rows['dn'] = match_rows.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            sel = st.selectbox("정확한 항목 선택", ["선택하세요"] + match_rows['dn'].tolist())
            
            if sel != "선택하세요":
                target_row = match_rows[match_rows['dn'] == sel].iloc[0]
                
                # 가이드북 열(Column) 순서대로 체크된 항목 수집
                # 순번, 분류, 개선내역 이후의 모든 열을 순회
                active_columns = []
                for col in df_g.columns:
                    col_name = str(col).strip()
                    if any(ex in col_name for ex in ["순번", "분류", "개선내역", "Unnamed"]):
                        continue
                    if is_checked(target_row[col]):
                        active_columns.append(col_name)

                # 파일별 분류 키워드
                r_keywords = ["일반현황", "점검사항", "자료생성", "자료수집기", "관제센터"]
                c_keywords = ["구조", "시료", "승인", "방법", "범위", "교정", "표준물질", "정도검사", "교정일자", "유량계", "누적값", "반복성", "드리프트", "재현성"]

                def get_matched_tabs(test_name, sheet_dict, f_type):
                    matched = []
                    tn_c = test_name.replace(" ", "")
                    for sn in sheet_dict.keys():
                        sn_c = str(sn).replace(" ", "")
                        # 1. 탭 이름에 가이드북 항목명이 포함되는지
                        if tn_c in sn_c or sn_c in tn_c:
                            matched.append(sn)
                        # 2. 외관 및 구조 등 포괄 규칙
                        elif f_type == "확인" and ("외관" in tn_c or "구조" in tn_c):
                            if any(k in sn_c for k in ["구조", "시료", "승인", "방법", "범위", "교정", "일자"]):
                                matched.append(sn)
                        elif f_type == "통합" and any(k in tn_c for k in ["점검사항", "자료생성"]):
                            if any(k in sn_c for k in ["점검", "생성", "전송", "관제"]):
                                matched.append(sn)
                    return list(dict.fromkeys(matched))

                c1, c2, c3 = st.columns(3)

                with c1:
                    st.header("1. 통합시험 조사표")
                    for item in active_columns:
                        if any(k in item for k in r_keywords):
                            st.write(f"**- {item}**") # 가이드북 순서대로 출력
                            tabs = get_matched_tabs(item, r_s, "통합")
                            for t in tabs:
                                with st.expander(f"└ {t} 탭 확인"):
                                    st.dataframe(r_s[t].fillna(""))

                with c2:
                    st.header("2. 확인검사 조사표")
                    for item in active_columns:
                        if any(k in item for k in c_keywords):
                            st.write(f"**- {item}**") # 가이드북 순서대로 출력
                            tabs = get_matched_tabs(item, c_s, "확인")
                            for t in tabs:
                                with st.expander(f"└ {t} 탭 확인"):
                                    st.dataframe(c_s[t].fillna(""))

                with c3:
                    st.header("3. 상대정확도 확인서")
                    for item in active_columns:
                        if "상대" in item:
                            st.write(f"**- {item}**") # 가이드북 순서대로 출력
                            for t in s_s.keys():
                                with st.expander(f"└ {t} 탭 확인"):
                                    st.dataframe(s_s[t].fillna(""))
