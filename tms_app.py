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

st.title("📋 수질 TMS 개선내역 매칭 결과")

if df_guide is not None:
    search_q = st.text_input("개선내역 입력", "")
    
    if search_q:
        match_rows = df_guide[df_guide.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        
        if not match_rows.empty:
            match_rows['dn'] = match_rows.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            sel = st.selectbox("항목 선택", ["선택하세요"] + match_rows['dn'].tolist())
            
            if sel != "선택하세요":
                row = match_rows[match_rows['dn'] == sel].iloc[0]
                
                # 가이드북에서 'ㅇ' 체크된 모든 시험명 추출
                active = [str(col).strip() for col in df_guide.columns if is_checked(row[col])]
                active = [t for t in active if not any(ex in t for ex in ["순번", "분류", "개선내역", "Unnamed"])]

                def find_matches(check_list, sheet_dict, f_type):
                    matched = []
                    cl_str = "".join(check_list).replace(" ", "")
                    for sn in sheet_dict.keys():
                        sn_c = str(sn).replace(" ", "")
                        # 1. 탭 이름 직접 매칭
                        if any(c.replace(" ", "") in sn_c or sn_c in c.replace(" ", "") for c in check_list):
                            matched.append(sn)
                        # 2. 확인검사 예외 규칙 (외관/구조 관련)
                        elif f_type == "확인" and any(k in cl_str for k in ["외관", "구조"]):
                            if any(k in sn_c for k in ["구조", "시료", "승인", "방법", "범위", "교정", "일자"]):
                                matched.append(sn)
                        # 3. 통합시험 예외 규칙
                        elif f_type == "통합" and any(k in cl_str for k in ["점검사항", "자료생성", "수집기"]):
                            if any(k in sn_c for k in ["점검", "생성", "전송", "관제"]):
                                matched.append(sn)
                    return list(set(matched))

                all_data = []
                c1, c2, c3 = st.columns(3)

                with c1:
                    st.header("1. 통합시험")
                    m_r = find_matches(active, r_sheets, "통합")
                    if m_r:
                        for m in m_r:
                            with st.expander(f"✅ {m}"):
                                st.dataframe(r_sheets[m].fillna(""))
                                t = r_sheets[m].copy(); t.insert(0, '시험분류', '통합시험'); t.insert(1, '탭이름', m)
                                all_data.append(t)
                    else: st.info("해당사항 없음")

                with c2:
                    st.header("2. 확인검사")
                    m_c = find_matches(active, c_sheets, "확인")
                    if m_c:
                        for m in m_c:
                            with st.expander(f"✅ {m}"):
                                st.dataframe(c_sheets[m].fillna(""))
                                t = c_sheets[m].copy(); t.insert(0, '시험분류', '확인검사'); t.insert(1, '탭이름', m)
                                all_data.append(t)
                    else: st.info("해당사항 없음")

                with c3:
                    st.header("3. 상대정확도")
                    if any("상대" in c for c in active):
                        for m in s_sheets.keys():
                            with st.expander(f"✅ {m}"):
                                st.dataframe(s_sheets[m].fillna(""))
                                t = s_sheets[m].copy(); t.insert(0, '시험분류', '상대정확도'); t.insert(1, '탭이름', m)
                                all_data.append(t)
                    else: st.info("해당사항 없음")

                if all_data:
                    out = BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
                        pd.concat(all_data).to_excel(wr, index=False)
                    st.download_button("📥 결과 보고서 다운로드", out.getvalue(), "TMS_Result.xlsx")
