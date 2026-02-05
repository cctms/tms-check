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
        
        # 1~3행을 모두 읽어서 진짜 '시험 이름'이 있는 행을 찾습니다.
        h_all = pd.read_excel(g_p, sheet_name=g_sn, nrows=3, header=None)
        
        # 보통 2행(index 1)이나 3행(index 2)에 "반복성", "제로드리프트" 같은 진짜 이름이 있습니다.
        # "수질TMS"라는 제목이 반복되는 행을 피하고 세부 항목이 있는 행을 선택합니다.
        raw_cols = h_all.iloc[1] # 일단 2행을 기준으로 시도
        if any("수질TMS" in str(x) for x in raw_cols): # 만약 2행도 제목이면 3행 선택
            raw_cols = h_all.iloc[2]
            
        new_cols = []
        for i, col in enumerate(raw_cols):
            c_str = str(col).strip()
            # Unnamed이거나 공백이면 앞/뒤 행에서 보완 (필요시)
            if "Unnamed" in c_str or c_str == "nan":
                new_cols.append(f"col_{i}")
            else:
                new_cols.append(c_str)
            
        df = pd.read_excel(g_p, sheet_name=g_sn, skiprows=2, header=None) # 데이터 시작 위치
        # 만약 skiprows=2가 제목을 포함한다면 숫자를 조정해야 함
        if any("수질TMS" in str(df.iloc[0, 0]) for _ in range(1)):
            df = pd.read_excel(g_p, sheet_name=g_sn, skiprows=3, header=None)
            
        df.columns = new_cols
        df.iloc[:, 1] = df.iloc[:, 1].ffill() # 분류 채우기
        
        return df, pd.read_excel(r_p, sheet_name=None) if r_p else {}, \
               pd.read_excel(c_p, sheet_name=None) if c_p else {}, \
               pd.read_excel(s_p, sheet_name=None) if s_p else {}, f_list
    except Exception as e:
        return None, None, None, None, [str(e)]

df, r_s, c_s, s_s, f_list = load_data()

def ck(v):
    if isinstance(v, pd.Series): v = v.iloc[0]
    if pd.isna(v): return False
    s = str(v).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

st.title("📋 수질 TMS 시험방법 자동 매칭")

if df is not None:
    q = st.text_input("개선내역 검색", "")
    if q:
        # 개선내역 열(3번째 열)에서 검색
        res = df[df.iloc[:, 2].astype(str).str.contains(q, na=False)].copy()
        if not res.empty:
            res['dn'] = res.apply(lambda x: f"[{x.iloc[1]}] {str(x.iloc[2]).strip()}", axis=1)
            sel = st.selectbox("항목선택", ["선택"] + res['dn'].tolist())
            
            if sel != "선택":
                row = res[res['dn'] == sel].iloc[0]
                
                # 체크된 항목 추출 로직 강화
                checked_items = []
                for i, val in enumerate(row):
                    if ck(val):
                        col_name = str(df.columns[i])
                        if not any(x in col_name for x in ["순번", "분류", "개선내역", "col_"]):
                            checked_items.append(col_name)

                st.success(f"📍 인식된 시험 종류: {', '.join(checked_items)}")
                
                all_d = []
                col1, col2, col3 = st.columns(3)

                # 매칭 로직 (예시로 주신 항목들 포함)
                def is_match(sheet_name, checked_list):
                    sn = sheet_name.replace(" ", "")
                    cl_str = "".join(checked_list).replace(" ", "")
                    
                    # 1. 직접 포함 관계
                    if any(c in sn or sn in c for c in checked_list): return True
                    
                    # 2. 통합시험 전송 관련 예외
                    if any(k in sn for k in ["일반현황", "점검사항", "자료생성", "전송", "관제센터"]):
                        if any(k in cl_str for k in ["일반현황", "점검사항", "자료생성", "전송"]): return True
                    
                    # 3. 확인검사 예외 (외관/구조 체크 시)
                    if any(k in cl_str for k in ["외관", "구조"]):
                        if any(k in sn for k in ["구조", "시료", "승인", "방법", "범위", "교정", "일자"]): return True
                        
                    # 4. 유량계/누적값 예외
                    if "유량" in cl_str and any(k in sn for k in ["유량", "누적"]): return True
                    
                    return False

                with col1:
                    st.subheader("1. 통합시험")
                    found = False
                    for s_n in r_s.keys():
                        if is_match(s_n, checked_items):
                            with st.expander(f"✅ {s_n}"):
                                st.dataframe(r_s[s_n].fillna(""))
                                t = r_s[s_n].copy(); t.insert(0, '시험', s_n); all_d.append(t)
                                found = True
                    if not found: st.info("해당사항 없음")

                with col2:
                    st.subheader("2. 확인검사")
                    found = False
                    for s_n in c_s.keys():
                        if is_match(s_n, checked_items):
                            with st.expander(f"✅ {s_n}"):
                                st.dataframe(c_s[s_n].fillna(""))
                                t = c_s[s_n].copy(); t.insert(0, '시험', s_n); all_d.append(t)
                                found = True
                    if not found: st.info("해당사항 없음")

                with col3:
                    st.subheader("3. 상대정확도")
                    if any("상대" in c for c in checked_items):
                        if s_s:
                            k = list(s_s.keys())[0]
                            with st.expander("✅ 상대정확도"):
                                st.dataframe(s_s[k].fillna(""))
                                t = s_s[k].copy(); t.insert(0, '시험', '상대정확도'); all_d.append(t)
                    else: st.info("해당사항 없음")

                if all_d:
                    out = BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
                        pd.concat(all_d).to_excel(wr, index=False)
                    st.download_button("📥 결과 다운로드", out.getvalue(), "TMS_Report.xlsx")
