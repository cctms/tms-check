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
        
        # 병합 헤더 처리
        h_df = pd.read_excel(g_p, sheet_name=g_sn, nrows=2, header=None)
        h_df.iloc[0] = h_df.iloc[0].ffill()
        new_cols = []
        for c1, c2 in zip(h_df.iloc[0], h_df.iloc[1]):
            c1_s, c2_s = str(c1) if pd.notna(c1) else "", str(c2) if pd.notna(c2) else ""
            name = f"{c1_s}_{c2_s}" if c1_s != c2_s and c2_s and "Unnamed" not in c2_s else c1_s
            new_cols.append(name.strip())
            
        df = pd.read_excel(g_p, sheet_name=g_sn, skiprows=2, header=None)
        df.columns = new_cols
        df.iloc[:, 1] = df.iloc[:, 1].ffill() 
        
        return df, pd.read_excel(r_p, sheet_name=None) if r_p else {}, \
               pd.read_excel(c_p, sheet_name=None) if c_p else {}, \
               pd.read_excel(s_p, sheet_name=None) if s_p else {}, f_list
    except Exception as e:
        return None, None, None, None, [str(e)]

df, r_s, c_s, s_s, f_list = load_data()

def ck(v):
    if pd.isna(v): return False
    return any(m in str(v).replace(" ", "").upper() for m in ['O', 'ㅇ', '○', 'V'])

st.title("📋 수질 TMS 시험항목 (교체 규칙 적용)")

if df is not None:
    q = st.text_input("개선내역 검색 (예: 기기교체)", "")
    if q:
        res = df[df.iloc[:, 2].astype(str).str.contains(q, na=False)].copy()
        if not res.empty:
            res['dn'] = res.apply(lambda x: f"[{x.iloc[1]}] {str(x.iloc[2]).strip()}", axis=1)
            sel = st.selectbox("항목선택", ["선택"] + res['dn'].tolist())
            
            if sel != "선택":
                row = res[res['dn'] == sel].iloc[0]
                is_replacement = "교체" in str(row.iloc[2]) # '교체' 키워드 확인
                
                # 1. 통합시험 필수 키워드 (교체 시)
                r_must = ["일반현황", "점검사항", "자료생성", "측정기기-자료수집기", "자료수집기-관제센터"]
                # 2. 확인검사 필수 키워드 (교체 시)
                c_must = ["구조", "시료채취", "형식승인", "측정방법", "측정범위", "교정기능", "표준물질", "정도검사", "교정일자", "유량계", "누적값"]

                all_d = []
                col1, col2, col3 = st.columns(3)

                with col1:
                    st.subheader("1. 통합시험")
                    f_r = False
                    for s_n in r_s.keys():
                        s_c = str(s_n).replace(" ", "")
                        # 교체면 필수항목이거나, 가이드북에 ㅇ가 있거나
                        if (is_replacement and any(k in s_c for k in r_must)) or \
                           any(ck(row[col]) and s_c in col.replace(" ", "") for col in df.columns):
                            with st.expander(f"✅ {s_n}"):
                                st.dataframe(r_s[s_n].fillna(""))
                                t = r_s[s_n].copy(); t.insert(0, '시험', s_n); all_d.append(t); f_r = True
                    if not f_r: st.info("해당사항 없음")

                with col2:
                    st.subheader("2. 확인검사")
                    f_c = False
                    for s_n in c_s.keys():
                        s_c = str(s_n).replace(" ", "")
                        # 교체 규칙 적용 (구조, 시료, 승인, 방법, 범위, 교정, 유량 등)
                        if (is_replacement and any(k in s_c for k in c_must)) or \
                           any(ck(row[col]) and s_c in col.replace(" ", "") for col in df.columns):
                            with st.expander(f"✅ {s_n}"):
                                st.dataframe(c_s[s_n].fillna(""))
                                t = c_s[s_n].copy(); t.insert(0, '시험', s_n); all_d.append(t); f_c = True
                    if not f_c: st.info("해당사항 없음")

                with col3:
                    st.subheader("3. 상대정확도")
                    f_s = False
                    if any("상대" in str(col) and ck(row[col]) for col in df.columns):
                        if s_s:
                            k = list(s_s.keys())[0]
                            with st.expander("✅ 상대정확도"):
                                st.dataframe(s_s[k].fillna(""))
                                t = s_s[k].copy(); t.insert(0, '시험', '상대정확도'); all_d.append(t); f_s = True
                    if not f_s: st.info("해당사항 없음")

                if all_d:
                    out = BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
                        pd.concat(all_d).to_excel(wr, index=False)
                    st.download_button("📥 결과 다운로드", out.getvalue(), "TMS_Report.xlsx")
