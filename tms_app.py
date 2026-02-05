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
        # 헤더를 찾기 위해 skiprows를 조정하거나 자동 감지 로직 추가 가능
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

                # 1. 통합시험 섹션
                with col1:
                    st.subheader("1. 통합시험")
                    found_r = False
                    # 통합시험은 기존 인덱스(3~10)를 유지하되 보수적으로 접근
                    t_l = [("1. 일반현황", 3), ("2. 하드웨어 규격", 4), ("3. 소프트웨어 기능 규격", 5), ("4. 자료정의", 6), ("5. 측정기기 점검사항", 7), ("6. 자료생성", 8), ("7. 측정기기-자료수집기", 9), ("8. 자료수집기-관제센터", 10)]
                    for nm, idx in t_l:
                        if idx < len(row) and (ck(row.iloc[idx]) or (is_c and idx in [9, 10])):
                            for s_name in r_s.keys():
                                if nm.replace(" ", "") in str(s_name).replace(" ", ""):
                                    with st.expander(f"✅ {nm}"):
                                        t = r_s[s_name].fillna(""); st.dataframe(t)
                                        t_exp = t.copy(); t_exp.insert(0, '시험', nm); all_d.append(t_exp)
                                        found_r = True
                    if not found_r: st.info("해당사항 없음")

                # 2. 확인검사 섹션 (강화된 로직)
                with col2:
                    st.subheader("2. 확인검사")
                    found_c = False
                    active_keywords = []
                    
                    # 확인검사 핵심 키워드 맵 (가이드북 열 이름 : 조사표 시트 키워드)
                    c_map = {
                        "외관": ["구조", "시료", "승인", "방법", "범위", "물질", "일자"],
                        "변동": ["전압"],
                        "절연": ["절연"],
                        "안정성": ["안정성"],
                        "반복성": ["반복성"],
                        "드리프트": ["제로", "스팬", "드리프트"],
                        "응답": ["응답"],
                        "직선성": ["직선성"],
                        "간섭": ["간섭"],
                        "검출": ["검출"],
                        "입지": ["입지조건"],
                        "유량": ["유량계", "누적값"]
                    }

                    # 가이드북의 모든 열을 돌면서 체크된 항목의 키워드를 active_keywords에 담음
                    for col_name in df.columns:
                        for key, sheets in c_map.items():
                            if key in str(col_name):
                                if ck(row[col_name]):
                                    active_keywords.extend(sheets)

                    if c_s and active_keywords:
                        # 2번 파일(확인검사)의 실제 시트 순서대로 확인
                        for s_name in c_s.keys():
                            s_clean = str(s_name).replace(" ", "")
                            if any(kw in s_clean for kw in active_keywords):
                                with st.expander(f"✅ {s_name}"):
                                    t = c_s[s_name].fillna(""); st.dataframe(t)
                                    t_exp = t.copy(); t_exp.insert(0, '시험', s_name); all_d.append(t_exp)
                                    found_c = True
                    
                    if not found_c: st.info("해당사항 없음")

                # 3. 상대정확도 섹션
                with col3:
                    st.subheader("3. 상대정확도")
                    found_s = False
                    for col_name in df.columns:
                        if "상대정확도" in str(col_name) and ck(row[col_name]):
                            if s_s:
                                k = list(s_s.keys())[0]
                                with st.expander("✅ 상대정확도"):
                                    t = s_s[k].fillna(""); st.dataframe(t)
                                    t_exp = t.copy(); t_exp.insert(0, '시험', '상대정확도'); all_d.append(t_exp)
                                    found_s = True
                    if not found_s: st.info("해당사항 없음")

                if all_d:
                    out = BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
                        pd.concat(all_d).to_excel(wr, index=False)
                    st.download_button("📥 전체 결과 다운로드", out.getvalue(), "TMS_Report.xlsx")
