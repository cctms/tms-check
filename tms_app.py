import streamlit as st
import pandas as pd
from io import BytesIO
import os
import re

st.set_page_config(page_title="TMS", layout="wide")

@st.cache_data
def load_data():
    try:
        f_list = os.listdir('.')
        g_p = next((f for f in f_list if '가이드북' in f or '시험방법' in f), None)
        r_p = next((f for f in f_list if '1.통합' in f), None)
        c_p = next((f for f in f_list if '2.확인' in f), None)
        s_p = next((f for f in f_list if '상대' in f or '3.' in f), None)
        if not g_p: return None, None, None, None, f_list
        
        xl_g = pd.ExcelFile(g_p)
        g_sn = next((s for s in xl_g.sheet_names if '가이드북' in s), xl_g.sheet_names[0])
        df = pd.read_excel(g_p, sheet_name=g_sn, skiprows=1)
        df.iloc[:, 1] = df.iloc[:, 1].ffill()
        
        # sheet_name=None으로 가져오면 순서가 유지된 딕셔너리가 반환됩니다.
        r_s = pd.read_excel(r_p, sheet_name=None) if r_p else {}
        c_s = pd.read_excel(c_p, sheet_name=None) if c_p else {}
        s_s = pd.read_excel(s_p, sheet_name=None) if s_p else {}
        return df, r_s, c_s, s_s, f_list
    except Exception as e:
        return None, None, None, None, [str(e)]

df, r_s, c_s, s_s, f_list = load_data()

def ck(v):
    if pd.isna(v): return False
    s = str(v).replace(" ", "").upper()
    return any(m in s for m in ['O', '○', 'V', 'CHECK'])

def find_sheet_strict(sheets_dict, target_name):
    if not sheets_dict: return None
    t_clean = str(target_name).replace(" ", "")
    for s_name in sheets_dict.keys():
        s_clean = str(s_name).replace(" ", "")
        if t_clean in s_clean or s_clean in t_clean: return s_name
    t_num = re.findall(r'\d+', str(target_name))
    if t_num:
        for s_name in sheets_dict.keys():
            s_num = re.findall(r'\d+', str(s_name))
            if s_num and t_num[0] == s_num[0]: return s_name
    return None

st.title("📋 수질 TMS 시험항목")

if df is not None:
    q = st.text_input("개선내역 검색 (예: 기기교체)", "")
    if q:
        res = df[df.iloc[:, 2].str.contains(q, na=False)].copy()
        if not res.empty:
            res['dn'] = res.apply(lambda x: f"[{x.iloc[1]}] {str(x.iloc[2]).strip()}", axis=1)
            sel = st.selectbox("항목선택", ["선택"] + res['dn'].tolist())
            if sel != "선택":
                row = res[res['dn'] == sel].iloc[0]
                is_c = "교체" in str(row.iloc[2])
                all_d = []
                col1, col2, col3 = st.columns(3)

                with col1:
                    st.subheader("1. 통합시험")
                    t_l = [("1. 일반현황", 3), ("2. 하드웨어 규격", 4), ("3. 소프트웨어 기능 규격", 5), ("4. 자료정의", 6), ("5. 측정기기 점검사항", 7), ("6. 자료생성", 8), ("7. 측정기기-자료수집기", 9), ("8. 자료수집기-관제센터", 10)]
                    for nm, idx in t_l:
                        if ck(row.iloc[idx]) or (is_c and idx in [9, 10]):
                            m_n = find_sheet_strict(r_s, nm)
                            if m_n:
                                with st.expander(f"✅ {nm}"):
                                    t = r_s[m_n].fillna(""); st.dataframe(t)
                                    t_exp = t.copy(); t_exp.insert(0, '시험', nm); all_d.append(t_exp)

                with col2:
                    st.subheader("2. 확인검사")
                    # 가이드북에서 'O' 표시된 항목들의 키워드 추출
                    c_guide = ["외관 및 구조", "전원전압 변동", "절연저항", "공급전압의 안정성", "반복성", "제로 및 스팬 드리프트", "응답시간", "직선성", "유입전류 안정성", "간섭영향", "검출한계"]
                    w_guide = ["구조", "시료", "승인", "방법", "범위", "물질", "일자"]
                    
                    active_keywords = []
                    for i, nm in enumerate(c_guide):
                        if ck(row.iloc[11+i]):
                            if nm == "외관 및 구조": active_keywords.extend(w_guide)
                            else: active_keywords.append(nm)
                    
                    # 실제 엑셀 시트 순서대로 순회하면서, 활성화된 키워드와 매칭되면 출력
                    if c_s:
                        for s_name in c_s.keys():
                            s_clean = str(s_name).replace(" ", "")
                            # 현재 시트가 가이드북에서 체크된 키워드 중 하나를 포함하는지 확인
                            if any(str(kw).replace(" ", "") in s_clean or s_clean in str(kw).replace(" ", "") for kw in active_keywords):
                                with st.expander(f"✅ {s_name}"):
                                    t = c_s[s_name].fillna(""); st.dataframe(t)
                                    t_exp = t.copy(); t_exp.insert(0, '시험', s_name); all_d.append(t_exp)

                with col3:
                    st.subheader("3. 상대정확도")
                    if ck(row.iloc[22]) and s_s:
                        k = list(s_s.keys())[0]
                        with st.expander("✅ 상대정확도"):
                            t = s_s[k].fillna(""); st.dataframe(t)
                            t_exp = t.copy(); t_exp.insert(0, '시험', '상대정확도'); all_d.append(t_exp)

                if all_d:
                    out = BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
                        pd.concat(all_d).to_excel(wr, index=False)
                    st.download_button("📥 전체 결과 다운로드", out.getvalue(), "TMS_Report.xlsx")
