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
        # 헤더가 있는 행을 자동으로 찾거나 지정 (보통 2번째 줄이 헤더이므로 skiprows=1)
        df = pd.read_excel(g_p, sheet_name=g_sn, skiprows=1)
        df.iloc[:, 1] = df.iloc[:, 1].ffill() # 분류 채우기
        
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
                all_d = []
                
                # 가이드북에서 'O' 표시된 모든 열의 이름을 수집
                checked_columns = []
                for col_name in df.columns:
                    if ck(row[col_name]):
                        checked_columns.append(str(col_name).replace(" ", ""))

                col1, col2, col3 = st.columns(3)

                # 1. 통합시험
                with col1:
                    st.subheader("1. 통합시험")
                    found_r = False
                    for s_name in r_s.keys():
                        s_clean = str(s_name).replace(" ", "")
                        # 체크된 열 이름이 시트 이름에 포함되어 있는지 확인
                        if any(col in s_clean or s_clean in col for col in checked_columns):
                            with st.expander(f"✅ {s_name}"):
                                t = r_s[s_name].fillna(""); st.dataframe(t)
                                t_exp = t.copy(); t_exp.insert(0, '시험', s_name); all_d.append(t_exp)
                                found_r = True
                    if not found_r: st.info("해당사항 없음")

                # 2. 확인검사 (탭 순서 유지)
                with col2:
                    st.subheader("2. 확인검사")
                    found_c = False
                    # 외관 및 구조 예외 처리 (구조, 시료, 승인 등 포함)
                    w_sub = ["구조", "시료", "승인", "방법", "범위", "물질", "일자"]
                    
                    if c_s:
                        for s_name in c_s.keys(): # 엑셀 시트 순서대로
                            s_clean = str(s_name).replace(" ", "")
                            
                            # 일반적인 매칭
                            match = any(col in s_clean or s_clean in col for col in checked_columns)
                            
                            # '외관' 관련 예외 매칭
                            if not match and any("외관" in col for col in checked_columns):
                                match = any(sub in s_clean for sub in w_sub)
                                
                            if match:
                                with st.expander(f"✅ {s_name}"):
                                    t = c_s[s_name].fillna(""); st.dataframe(t)
                                    t_exp = t.copy(); t_exp.insert(0, '시험', s_name); all_d.append(t_exp)
                                    found_c = True
                    if not found_c: st.info("해당사항 없음")

                # 3. 상대정확도
                with col3:
                    st.subheader("3. 상대정확도")
                    found_s = False
                    if any("상대정확도" in col for col in checked_columns):
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
