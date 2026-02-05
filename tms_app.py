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
        
        # 헤더 처리: 병합 셀 대응
        h_df = pd.read_excel(g_p, sheet_name=g_sn, nrows=2, header=None)
        h_df.iloc[0] = h_df.iloc[0].ffill()
        new_cols = []
        for i in range(len(h_df.columns)):
            c1 = str(h_df.iloc[0, i]) if pd.notna(h_df.iloc[0, i]) else ""
            c2 = str(h_df.iloc[1, i]) if pd.notna(h_df.iloc[1, i]) else ""
            # 대분류_소분류 형태 (Unnamed 방지)
            name = f"{c1}_{c2}" if c1 != c2 and c2 and "Unnamed" not in c2 else c1
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

# ㅇ, O, ○ 등이 포함된 모든 텍스트를 체크로 인식
def ck(v):
    if isinstance(v, pd.Series): v = v.iloc[0]
    if pd.isna(v): return False
    s = str(v).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎'])

st.title("📋 수질 TMS 시험방법 (2025 최종)")

if df is not None:
    q = st.text_input("개선내역 검색", "")
    if q:
        res = df[df.iloc[:, 2].astype(str).str.contains(q, na=False)].copy()
        if not res.empty:
            res['dn'] = res.apply(lambda x: f"[{x.iloc[1]}] {str(x.iloc[2]).strip()}", axis=1)
            sel = st.selectbox("항목선택", ["선택"] + res['dn'].tolist())
            
            if sel != "선택":
                row = res[res['dn'] == sel].iloc[0]
                
                # 1. 가이드북에서 'ㅇ'가 포함된 열(시험종류)을 모두 추출
                checked_items = []
                for i, val in enumerate(row):
                    if ck(val):
                        col_name = df.columns[i]
                        # '통합시험_일반현황' 같은 이름에서 '일반현황'만 추출
                        clean_name = col_name.split('_')[-1] if '_' in col_name else col_name
                        if clean_name not in ["순번", "분류", "개선내역"] and "Unnamed" not in clean_name:
                            checked_items.append(clean_name.replace(" ", ""))

                if checked_items:
                    st.info(f"📍 체크된 항목: {', '.join(checked_items)}")
                
                all_d = []
                col1, col2, col3 = st.columns(3)

                # 2. 체크된 항목과 시트 이름 매칭 (예외 규칙 포함)
                def is_match(sheet_name, checked_list):
                    s_n = sheet_name.replace(" ", "")
                    # 직접 매칭
                    if any(c in s_n or s_n in c for c in checked_list):
                        return True
                    # '외관 및 구조' 체크 시 '시료채취조', '형식승인' 등 포함
                    if "외관" in "".join(checked_list) or "구조" in "".join(checked_list):
                        if any(k in s_n for k in ["구조", "시료", "승인", "방법", "범위", "일자"]):
                            return True
                    # '유량' 관련 매칭
                    if "유량" in "".join(checked_list) and any(k in s_n for k in ["유량", "누적"]):
                        return True
                    return False

                with col1:
                    st.subheader("1. 통합시험")
                    f_r = False
                    for s_n in r_s.keys():
                        if is_match(s_n, checked_items):
                            with st.expander(f"✅ {s_n}"):
                                st.dataframe(r_s[s_n].fillna(""))
                                t = r_s[s_n].copy(); t.insert(0, '시험', s_n); all_d.append(t); f_r = True
                    if not f_r: st.info("해당사항 없음")

                with col2:
                    st.subheader("2. 확인검사")
                    f_c = False
                    for s_n in c_s.keys():
                        if is_match(s_n, checked_items):
                            with st.expander(f"✅ {s_n}"):
                                st.dataframe(c_s[s_n].fillna(""))
                                t = c_s[s_n].copy(); t.insert(0, '시험', s_n); all_d.append(t); f_c = True
                    if not f_c: st.info("해당사항 없음")

                with col3:
                    st.subheader("3. 상대정확도")
                    f_s = False
                    if any("상대" in c for c in checked_items):
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
