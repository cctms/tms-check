import streamlit as st
import pandas as pd
from io import BytesIO
import os

st.set_page_config(page_title="TMS", layout="wide")

# 1. 파일 로드 함수 (이름 매칭 강화)
@st.cache_data
def load_data():
    try:
        f_list = os.listdir('.')
        # 파일명에 특정 단어가 포함되어 있는지 확인
        g_p = next((f for f in f_list if '가이드북' in f or '시험방법' in f), None)
        r_p = next((f for f in f_list if '1.통합' in f), None)
        c_p = next((f for f in f_list if '2.확인' in f), None)
        s_p = next((f for f in f_list if '상대' in f or '3.' in f), None)
        
        if not g_p: 
            return None, None, None, None, f_list
            
        df = pd.read_excel(g_p, sheet_name='★최종(가이드북)', skiprows=1)
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

# 2. 파일이 없을 경우 경고 메시지 표시
if df is None:
    st.error("⚠️ '가이드북' 엑셀 파일을 찾을 수 없습니다.")
    st.info(f"현재 폴더의 파일 목록: {f_list}")
    st.write("파일명에 '가이드북'이라는 단어가 포함되어 있는지 확인해 주세요.")

# 3. 검색창은 파일 유무와 상관없이 표시 (단, 데이터가 있어야 작동)
q = st.text_input("개선내역 검색 (예: 기기교체)", "")

if q and df is not None:
    res = df[df.iloc[:, 2].str.contains(q, na=False)].copy()
    if not res.empty:
        res['dn'] = res.apply(lambda x: f"[{x.iloc[1]}] {str(x.iloc[2]).strip()}", axis=1)
        sel = st.selectbox("항목선택", ["선택"] + res['dn'].tolist())
        
        if sel != "선택":
            row = res[res['dn'] == sel].iloc[0]
            txt = str(row.iloc[2])
            is_c = "교체" in txt
            all_d = []
            c1, c2, c3 = st.columns(3)

            with c1:
                st.subheader("1. 통합시험")
                t_l = [("1. 일반현황", 3), ("2. 하드웨어 규격", 4), ("3. 소프트웨어 기능 규격", 5), ("4. 자료정의", 6), ("5. 측정기기 점검사항", 7), ("6. 자료생성", 8), ("7. 측정기기-자료수집기", 9), ("8. 자료수집기-관제센터", 10)]
                for nm, idx in t_l:
                    if ck(row.iloc[idx]) or (is_c and idx in [9, 10]):
                        # 시트명에 해당 이름이 포함되어 있는지 확인 (부분 일치 허용)
                        m = next((s for s in r_s.keys() if nm.strip() in s.strip()), None)
                        if m:
                            with st.expander(nm):
                                t = r_s[m].fillna(""); st.dataframe(t)
                                t_exp = t.copy(); t_exp.insert(0, '시험', nm); all_d.append(t_exp)
                        else:
                            st.warning(f"⚠️ {nm} 시트를 찾을 수 없음")

            with c2:
                st.subheader("2. 확인검사")
                c_l = ["외관 및 구조", "전원전압 변동", "절연저항", "공급전압의 안정성", "반복성", "제로 및 스팬 드리프트", "응답시간", "직선성", "유입전류 안정성", "간섭영향", "검출한계"]
                w_l = ["측정소 구조 및 설비", "시료채취조", "형식승인", "측정방법", "측정범위", "교정기능(표준물질)", "정도검사 교정일자"]
                for i, nm in enumerate(c_l):
                    if ck(row.iloc[11+i]):
                        if nm == "외관 및 구조":
                            for wn in w_l:
                                if wn in c_s:
                                    with st.expander(wn):
                                        t = c_s[wn].fillna(""); st.dataframe(t)
                                        t_exp = t.copy(); t_exp.insert(0, '시험', wn); all_d.append(t_exp)
                        elif nm in c_s:
                            with st.expander(nm):
                                t = c_s[nm].fillna(""); st.dataframe(t)
                                t_exp = t.copy(); t_exp.insert(0, '시험', nm); all_d.append(t_exp)

            with c3:
                st.subheader("3. 상대정확도")
                if ck(row.iloc[22]):
                    if s_s:
                        k = list(s_s.keys())[0]
                        with st.expander("상대정확도 결과서"):
                            t = s_s[k].fillna(""); st.dataframe(t)
                            t_exp = t.copy(); t_exp.insert(0, '시험', '상대정확도'); all_d.append(t_exp)

            if all_d:
                out = BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
                    pd.concat(all_d).to_excel(wr, index=False)
                st.download_button("📥 전체 결과 엑셀 다운로드", out.getvalue(), "TMS_Report.xlsx")
    else:
        st.warning("검색 결과가 없습니다.")
