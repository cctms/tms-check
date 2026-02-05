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
        
        # 가이드북 헤더 탐색 로직 (실제 시험명이 있는 행 찾기)
        guide_raw = pd.read_excel(g_p, header=None)
        header_idx = 2
        for i in range(min(5, len(guide_raw))):
            row_str = "".join(guide_raw.iloc[i].astype(str))
            if any(k in row_str for k in ["반복성", "제로드리프트", "일반현황"]):
                header_idx = i
                break
        
        df_guide = pd.read_excel(g_p, skiprows=header_idx)
        df_guide.iloc[:, 1] = df_guide.iloc[:, 1].ffill()
        
        r_sheets = pd.read_excel(r_p, sheet_name=None) if r_p else {}
        c_sheets = pd.read_excel(c_p, sheet_name=None) if c_p else {}
        s_sheets = pd.read_excel(s_p, sheet_name=None) if s_p else {}
        
        return df_guide, r_sheets, c_sheets, s_sheets
    except Exception as e:
        return None, None, None, None

df_guide, r_sheets, c_sheets, s_sheets = load_all_data()

def is_checked(value):
    if pd.isna(value): return False
    val = str(value).replace(" ", "").upper()
    return any(m in val for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

st.title("📋 개선내역별 수행항목 매칭 시스템")

if df_guide is not None:
    search_q = st.text_input("개선내역 입력 (예: 기기교체)", "")
    
    if search_q:
        match_rows = df_guide[df_guide.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        
        if not match_rows.empty:
            match_rows['display_name'] = match_rows.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            selected_item = st.selectbox("항목 선택", ["선택하세요"] + match_rows['display_name'].tolist())
            
            if selected_item != "선택하세요":
                target_row = match_rows[match_rows['display_name'] == selected_item].iloc[0]
                
                # 가이드북 체크 항목 추출
                active_tests = [str(col).strip() for col in df_guide.columns if is_checked(target_row[col])]
                active_tests = [t for t in active_tests if not any(ex in t for ex in ["순번", "분류", "개선내역", "Unnamed"])]

                # --- 여기서부터 파일별로 분류하여 표출 ---
                st.write("### 🔍 가이드북 기준 시험 분류")
                
                # 분류 기준 키워드
                r_must = ["일반현황", "점검사항", "자료생성", "자료수집기", "관제센터"]
                c_must = ["구조", "시료", "승인", "방법", "범위", "교정", "표준물질", "정도검사", "교정일자", "유량계", "누적값", "반복성", "드리프트", "재현성"]
                
                # 1. 통합시험 분류 항목
                r_list = [t for t in active_tests if any(k in t for k in r_must)]
                # 2. 확인검사 분류 항목
                c_list = [t for t in active_tests if any(k in t for k in c_must)]
                # 3. 상대정확도 분류 항목
                s_list = [t for t in active_tests if "상대" in t]

                m_col1, m_col2, m_col3 = st.columns(3)
                with m_col1:
                    st.info(f"**[통합시험]**\n\n" + ("\n".join([f"- {i}" for i in r_list]) if r_list else "해당 없음"))
                with m_col2:
                    st.success(f"**[확인검사]**\n\n" + ("\n".join([f"- {i}" for i in c_list]) if c_list else "해당 없음"))
                with m_col3:
                    st.warning(f"**[상대정확도]**\n\n" + ("\n".join([f"- {i}" for i in s_list]) if s_list else "해당 없음"))
                
                st.write("---")

                # --- 실제 탭 데이터 매칭 ---
                def find_matches(check_list, sheet_dict, file_type):
                    matched = []
                    cl_str = "".join(check_list).replace(" ", "")
                    for sn in sheet_dict.keys():
                        sn_clean = str(sn).replace(" ", "")
                        # 직접 매칭 또는 포괄 매칭
                        if any(c.replace(" ", "") in sn_clean or sn_clean in c.replace(" ", "") for c in check_list):
                            matched.append(sn)
                        elif file_type == "확인" and ("외관" in cl_str or "구조" in cl_str):
                            if any(k in sn_clean for k in ["구조", "시료", "승인", "방법", "범위", "교정", "일자"]):
                                matched.append(sn)
                    return list(set(matched))

                c1, c2, c3 = st.columns(3)
                all_data = []

                with c1:
                    st.subheader("1. 통합시험 조사표")
                    matches = find_matches(r_list, r_sheets, "통합")
                    for m in matches:
                        with st.expander(f"✅ {m}"):
                            st.dataframe(r_sheets[m].fillna(""))
                            t = r_sheets[m].copy(); t.insert(0, '탭이름', m); all_data.append(t)

                with c2:
                    st.subheader("2. 확인검사 조사표")
                    matches = find_matches(c_list, c_sheets, "확인")
                    for m in matches:
                        with st.expander(f"✅ {m}"):
                            st.dataframe(c_sheets[m].fillna(""))
                            t = c_sheets[m].copy(); t.insert(0, '탭이름', m); all_data.append(t)

                with c3:
                    st.subheader("3. 상대정확도 확인서")
                    if s_list:
                        for m in s_sheets.keys():
                            with st.expander(f"✅ {m}"):
                                st.dataframe(s_sheets[m].fillna(""))
                                t = s_sheets[m].copy(); t.insert(0, '탭이름', m); all_data.append(t)

                if all_data:
                    out = BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
                        pd.concat(all_data).to_excel(wr, index=False)
                    st.download_button("📥 통합 결과 다운로드", out.getvalue(), "TMS_Matching_Result.xlsx")
