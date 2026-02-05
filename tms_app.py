import streamlit as st
import pandas as pd
from io import BytesIO
import os

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
        g_sn = next((s for s in xl_g.sheet_names if '가이드북' in s or '시험방법' in s), xl_g.sheet_names[0])
        df = pd.read_excel(g_p, sheet_name=g_sn, skiprows=1)
        df.iloc[:, 1] = df.iloc[:, 1].ffill() 
        
        r_s = pd.read_excel(r_p, sheet_name=None) if r_p else {}
        c_s = pd.read_excel(c_p, sheet_name=None) if c_p else {}
        s_s = pd.read_excel(s_p, sheet_name=None) if s_p else {}
        return df, r_s, c_s, s_s, f_list
    except Exception as e:
        return None, None, None, None, [str(e)]

df, r_s, c_s, s_s, f_list = load_data()

# ㅇ, O, ○, V 등 어떤 문자라도 포함되어 있으면 체크로 간주하는 함수
def ck(v):
    if pd.isna(v): return False
    s = str(v).replace(" ", "").upper()
    # 한국어 'ㅇ'과 영문 'O', 동그라미 기호 등을 모두 포함
    check_marks = ['O', 'ㅇ', '○', '◎', 'V', 'CHECK']
    return any(m in s for m in check_marks)

st.title("📋 수질 TMS 시험항목 (2025 최종 기준)")

if df is not None:
    q = st.text_input("개선내역 검색 (예: 기기교체)", "")
    if q:
        res = df[df.iloc[:, 2].str.contains(q, na=False)].copy()
        if not res.empty:
            res['dn'] = res.apply(lambda x: f"[{x.iloc[1]}] {str(x.iloc[2]).strip()}", axis=1)
            sel = st.selectbox("항목선택", ["선택"] + res['dn'].tolist())
            
            if sel != "선택":
                row = res[res['dn'] == sel].iloc[0]
                
                # 'ㅇ'이 포함된 열 이름(시험종류) 추출
                checked_columns = []
                for col_name in df.columns:
                    if ck(row[col_name]):
                        checked_columns.append(str(col_name).strip())

                # 상단에 선택된 시험 종류 표시
                if checked_keywords := [c for c in checked_columns if c not in ["순번", "분류", "개선내역"]]:
                    st.success(f"🔍 **판단된 시험 종류:** {', '.join(checked_keywords)}")
                
                all_d = []
                col1, col2, col3 = st.columns(3)

                # 1. 통합시험
                with col1:
                    st.subheader("1. 통합시험")
                    found_r = False
                    for s_name in r_s.keys():
                        s_clean = str(s_name).replace(" ", "")
                        if any(kw.replace(" ", "") in s_clean or s_clean in kw.replace(" ", "") for kw in checked_columns):
                            with st.expander(f"✅ {s_name}"):
                                t = r_s[s_name].fillna(""); st.dataframe(t)
                                t_exp = t.copy(); t_exp.insert(0, '시험', s_name); all_d.append(t_exp)
                                found_r = True
                    if not found_r: st.info("해당사항 없음")

                # 2. 확인검사 (입지조건, 유량계 포함)
                with col2:
                    st.subheader("2. 확인검사")
                    found_c = False
                    w_sub = ["구조", "시료", "승인", "방법", "범위", "물질", "일자"]
                    flow_keywords = ["유량", "누적"]
                    
                    if c_s:
                        for s_name in c_s.keys():
                            s_clean = str(s_name).replace(" ", "")
                            match = False
                            
                            # 1) 열 이름 매칭
                            if any(kw.replace(" ", "") in s_clean or s_clean in kw.replace(" ", "") for kw in checked_columns):
                                match = True
                            # 2) 외관 및 구조 예외
                            if not match and any("외관" in kw for kw in checked_columns):
                                if any(sub in s_clean for sub in w_sub): match = True
                            # 3) 유량계/누적값 예외
                            if not match and any(f_kw in "".join(checked_columns) for f_kw in flow_keywords):
                                if any(f_kw in s_clean for f_kw in flow_keywords): match = True
                                    
                            if match:
                                with st.expander(f"✅ {s_name}"):
                                    t = c_s[s_name].fillna(""); st.dataframe(t)
                                    t_exp = t.copy(); t_exp.insert(0, '시험', s_name); all_d.append(t_exp)
                                    found_c = True
                    if not found_c: st.info("해당사항 없음")

                # 3. 상대정확도
                with col3:
                    st.subheader("3. 상대정확도")
                    if any("상대" in kw for kw in checked_columns):
                        if s_s:
                            k = list(s_s.keys())[0]
                            with st.expander("✅ 상대정확도"):
                                t = s_s[k].fillna(""); st.dataframe(t)
                                t_exp = t.copy(); t_exp.insert(0, '시험', '상대정확도'); all_d.append(t_exp)
                    else: st.info("해당사항 없음")

                if all_d:
                    out = BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
                        pd.concat(all_d).to_excel(wr, index=False)
                    st.download_button("📥 전체 결과 다운로드", out.getvalue(), "TMS_Report.xlsx")
