import streamlit as st
import pandas as pd
import os

# 페이지 설정
st.set_page_config(page_title="수질 TMS 스마트 가이드", layout="wide")

# 디자인 CSS
st.markdown("""
    <style>
    .main-title { font-size: 2.2rem; font-weight: 800; color: #1E3A8A; text-align: center; margin-bottom: 2rem; }
    .section-header { background: #1E3A8A; color: white; padding: 10px; border-radius: 5px; text-align: center; font-weight: 600; margin-bottom: 15px; }
    </style>
    """, unsafe_allow_html=True)

@st.cache_data
def load_all_resources():
    try:
        f_list = os.listdir('.')
        g_p = next((f for f in f_list if '가이드북' in f or '시험방법' in f), None)
        r_p = next((f for f in f_list if '1.통합' in f), None)
        c_p = next((f for f in f_list if '2.확인' in f), None)
        s_p = next((f for f in f_list if '상대' in f or '3.' in f), None)
        
        if not g_p: return None, None, None, None, None
        
        # 1. 가이드북 구조 분석
        df_raw = pd.read_excel(g_p, header=None)
        h_idx = 0
        for i in range(len(df_raw)):
            row_vals = df_raw.iloc[i].astype(str).values
            if "통합시험" in row_vals and "확인검사" in row_vals:
                h_idx = i
                break
        
        top_h = df_raw.iloc[h_idx].ffill() 
        sub_h = df_raw.iloc[h_idx + 1]     
        data_df = df_raw.iloc[h_idx + 2:].reset_index(drop=True)
        data_df.iloc[:, 1] = data_df.iloc[:, 1].ffill()
        
        # 2. 상세 조사표 데이터 로드 (딕셔너리 형태)
        r_data = pd.read_excel(r_p, sheet_name=None) if r_p else {}
        c_data = pd.read_excel(c_p, sheet_name=None) if c_p else {}
        s_data = pd.read_excel(s_p, sheet_name=None) if s_p else {}
        
        return data_df, top_h, sub_h, {"통합": r_data, "확인": c_data, "상대": s_data}
    except:
        return None, None, None, None

df, top_h, sub_h, survey_data = load_all_resources()

def is_ok(val):
    s = str(val).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

st.markdown('<p class="main-title">🌊 수질 TMS 수행항목 & 상세조사표</p>', unsafe_allow_html=True)

if df is not None:
    c_left, c_mid, c_right = st.columns([1, 2, 1])
    with c_mid:
        search_q = st.text_input("🔍 개선내역 키워드 입력", placeholder="예: 측정기기 교체")
    
    if search_q:
        matches = df[df.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        if not matches.empty:
            matches['dp'] = matches.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            with c_mid:
                sel = st.selectbox("📌 항목 선택", ["선택하세요"] + matches['dp'].tolist())
            
            if sel != "선택하세요":
                target_row = matches[matches['dp'] == sel].iloc[0]
                st.write("---")
                col1, col2, col3 = st.columns(3)
                
                col1.markdown('<p class="section-header">🛠 1. 통합시험</p>', unsafe_allow_html=True)
                col2.markdown('<p class="section-header">⚖️ 2. 확인검사</p>', unsafe_allow_html=True)
                col3.markdown('<p class="section-header">📊 3. 상대정확도</p>', unsafe_allow_html=True)

                for i in range(3, len(df.columns)):
                    if is_ok(target_row[i]):
                        cat, name = str(top_h[i]), str(sub_h[i])
                        
                        # 출력할 위치 결정
                        target_col = col1 if "통합" in cat else col2 if "확인" in cat else col3
                        
                        # 펼침(Expander) 구성
                        with target_col:
                            with st.expander(f"✅ {name}"):
                                # 해당 시험명과 유사한 이름의 시트 찾기
                                found_data = False
                                current_group = "통합" if "통합" in cat else "확인" if "확인" in cat else "상대"
                                sheets = survey_data.get(current_group, {})
                                
                                for s_name, s_df in sheets.items():
                                    if s_name.replace(" ","") in name.replace(" ","") or name.replace(" ","") in s_name.replace(" ",""):
                                        st.dataframe(s_df.fillna(""), use_container_width=True)
                                        found_data = True
                                        break
                                
                                if not found_data:
                                    st.info("해당 시험의 상세 조사표 시트를 찾을 수 없습니다.")

else:
    st.error("가이드북 및 조사표 파일을 확인해주세요.")
