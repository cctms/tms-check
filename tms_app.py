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
                            # 통합시험 매칭 로직
                            for s_name in r_s.keys():
                                if nm.replace(" ", "") in str(s_name).replace(" ", ""):
                                    with st.expander(f"✅ {nm}"):
                                        t = r_s[s_name].fillna(""); st.dataframe(t)
                                        t_exp = t.copy(); t_exp.insert(0, '시험', nm); all_d.append(t_exp)

                with col2:
                    st.subheader("2. 확인검사")
                    # 가이드북 컬럼 순서에 따른 매칭 키워드 (인덱스 11번부터 시작)
                    # 입지조건, 유량계 등을 가이드북의 열 순서에 맞춰 리스트업했습니다.
                    c_guide = [
                        ("외관 및 구조", 11), ("전원전압 변동", 12), ("절연저항", 13), 
                        ("공급전압의 안정성", 14), ("반복성", 15), ("제로 및 스팬 드리프트", 16), 
                        ("응답시간", 17), ("직선성", 18), ("유입전류 안정성", 19), 
                        ("간섭영향", 20), ("검출한계", 21), 
                        ("입지조건", None), ("유량계", None) # 가이드북에 별도 열이 있다면 인덱스 추가 필요
                    ]
                    
                    # 만약 가이드북 엑셀에 '입지조건'이나 '유량계' 열이 별도로 있다면 
                    # 아래 active_keywords에 추가되는 로직이 작동합니다.
                    w_sub = ["구조", "시료", "승인", "방법", "범위", "물질", "일자"]
                    active_keywords = []

                    for nm, idx in c_guide:
                        # 인덱스가 지정된 경우 해당 열의 체크 여부 확인
                        if idx is not None and ck(row.iloc[idx]):
                            if nm == "외관 및 구조": active_keywords.extend(w_sub)
                            else: active_keywords.append(nm)
                        # 만약 명칭으로 가이드북 열을 찾아야 한다면 (예: 22번 이후 열에 입지조건 등이 있는 경우)
                        elif idx is None:
                            # 가이드북 행 전체에서 해당 명칭이 체크되었는지 확인하는 로직 (필요시)
                            for col_idx, val in enumerate(row):
                                if nm in str(df.columns[col_idx]) and ck(val):
                                    active_keywords.append(nm)
                    
                    if c_s:
                        for s_name in c_s.keys():
                            s_clean = str(s_name).replace(" ", "")
                            if any(str(kw).replace(" ", "") in s_clean for kw in active_keywords):
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
