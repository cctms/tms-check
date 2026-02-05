import streamlit as st
import pandas as pd
from io import BytesIO
import os

st.set_page_config(page_title="TMS 도구", layout="wide")
st.title("📋 수질 TMS 시험항목")

@st.cache_data
def load_data():
    try:
        f_list = os.listdir('.')
        g_p = next((f for f in f_list if '가이드북' in f), None)
        r_p = next((f for f in f_list if '1.통합' in f), None)
        c_p = next((f for f in f_list if '2.확인' in f), None)
        s_p = next((f for f in f_list if '상대' in f), None)
        if not g_p: return None, None, None, None
        df = pd.read_excel(g_p, sheet_name='★최종(가이드북)', skiprows=1)
        df.iloc[:, 1] = df.iloc[:, 1].ffill()
        r_s = pd.read_excel(r_p, sheet_name=None) if r_p else {}
        c_s = pd.read_excel(c_p, sheet_name=None) if c_p else {}
        s_s = pd.read_excel(s_p, sheet_name=None) if s_p else {}
        return df, r_s, c_s, s_s
    except: return None, None, None, None

df, r_s, c_s, s_s = load_data()

def ck(v):
    if pd.isna(v): return False
    s = str(v).replace(" ", "").upper()
    return any(m in s for m in ['O', '○', 'V', 'CHECK'])

if df is not None:
    q = st.text_input("개선내역 검색 (예: 기기교체)", "")
    if q:
        res = df[df.iloc[:, 2].str.contains(q, na=False)].copy()
        if not res.empty:
            res['dn'] = res.apply(lambda x: f"[{x.iloc[1]}] {str(x.iloc[2]).strip()}", axis=1)
            sel = st.selectbox("검색결과", ["선택"] + res['dn'].tolist())
            if sel != "선택":
                row = res[res['dn'] == sel].iloc[0]
                txt = str(row.iloc[2])
                all_d = []
                col1, col2, col3 = st.columns(3)

                with col1:
                    st.subheader("1. 통합시험")
                    t_list = [("1. 일반현황", 3), ("2. 하드웨어 규격", 4), ("3. 소프트웨어 기능 규격", 5), ("4. 자료정의", 6), ("5. 측정기기 점검사항", 7), ("6. 자료생성", 8), ("7. 측정기기-자료수집기", 9), ("8. 자료수집기-관제센터", 10)]
                    is_c = "교체" in txt
                    for nm, idx in t_list:
                        if ck(row.iloc[idx]) or (is_c and idx in [9, 10]):
                            m_n = next((s for s in r_s.keys() if nm in s), None)
                            if m_n:
                                with st.expander(f"✅ {nm}"):
                                    tmp = r_s[m_n].fillna(""); st.dataframe(tmp)
                                    tmp.insert(0, '시험', nm); all_d.append(tmp)

                with col2:
                    st.subheader("2. 확인검사")
                    # 에러 났던 부분을 아주 짧은 리스트로 대체
                    c_l = ["외관 및 구조", "전원전압 변동", "절연저항", "공급전압의 안정성", "반복성", "제로 및 스팬 드리프트", "응답시간", "직선성", "유입전류 안정성", "간섭영향", "검출한계"]
                    w_l = ["측정소 구조 및 설비", "시료채취조", "형식승인", "측정방법", "측정범위", "교정기능(표준물질)", "정도검사 교정일자"]
                    for i, nm in enumerate(c_l):
                        if ck(row.iloc[11+i]):
                            if nm == "외관 및 구조":
                                for wn in w_l:
                                    if wn in c_s:
                                        with st.expander(f"✅ {wn}"):
                                            tmp = c_s[wn].fillna(""); st.dataframe(tmp)
                                            tmp.insert(0, '시험', wn); all_d.append(tmp)
                            elif nm in c_s:
                                with st.expander(f"✅ {nm}"):
                                    tmp = c_s[nm].fillna(""); st.dataframe(tmp)
                                    tmp.insert(0, '시험', nm); all_d.append(tmp)

                with col3:
                    st.subheader("3. 상대정확도")
                    if ck(row.iloc[22]):
                        if s_s:
                            k = list(s_s.keys())[0]
                            with st.expander("✅ 상대정확도"):
                                tmp = s_s[k].fillna(""); st.dataframe(tmp)
                                tmp.insert(0, '시험', '상대정확도'); all_d.append(tmp)

                if all_d:
                    out = BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
                        pd.concat(all_d).to_excel(wr, index=False)
                    st.download_button("📥 엑셀 다운로드", out.getvalue(), "TMS_Report.xlsx")
