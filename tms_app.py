import streamlit as st
import pandas as pd
import os

# 페이지 설정 (와이드 모드 및 타이틀)
st.set_page_config(page_title="수질 TMS 스마트 가이드", layout="wide")

# 커스텀 CSS로 디자인 입히기
st.markdown("""
    <style>
    .test-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 15px;
        border-left: 5px solid #007bff;
        margin-bottom: 10px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
    }
    .main-title {
        font-size: 2.5rem;
        font-weight: 800;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        background: #1E3A8A;
        color: white;
        padding: 10px;
        border-radius: 5px;
        text-align: center;
        font-weight: 600;
        margin-bottom: 15px;
    }
    </style>
    """, unsafe_allow_html=True)

@st.cache_data
def load_data():
    try:
        f_list = os.listdir('.')
        g_p = next((f for f in f_list if '가이드북' in f or '시험방법' in f), None)
        if not g_p: return None, None, None
        
        df_raw = pd.read_excel(g_p, header=None)
        h_idx = 0
        for i in range(len(df_raw)):
            row_vals = df_raw.iloc[i].astype(str).values
            if any("통합시험" in v for v in row_vals) and any("확인검사" in v for v in row_vals):
                h_idx = i
                break
        
        top_h = df_raw.iloc[h_idx].ffill() 
        sub_h = df_raw.iloc[h_idx + 1]     
        data_df = df_raw.iloc[h_idx + 2:].reset_index(drop=True)
        data_df.iloc[:, 1] = data_df.iloc[:, 1].ffill()
        
        return data_df, top_h, sub_h
    except:
        return None, None, None

df, top_h, sub_h = load_data()

def is_ok(val):
    s = str(val).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

# 제목 섹션
st.markdown('<p class="main-title">🌊 수질 TMS 수행항목 가이드</p>', unsafe_allow_html=True)

if df is not None:
    # 검색창 디자인
    with st.container():
        c_left, c_mid, c_right = st.columns([1, 2, 1])
        with c_mid:
            search_q = st.text_input("🔍 개선내역 키워드 입력", placeholder="예: 측정기기 교체")
    
    if search_q:
        matches = df[df.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        
        if not matches.empty:
            matches['dp'] = matches.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            with c_mid:
                sel = st.selectbox("📌 정확한 항목을 선택하세요", ["선택하세요"] + matches['dp'].tolist())
            
            if sel != "선택하세요":
                target_row = matches[matches['dp'] == sel].iloc[0]
                
                st.markdown("---")
                
                # 3단 컬럼 배치
                col1, col2, col3 = st.columns(3)
                
                # 섹션별 헤더 디자인
                col1.markdown('<p class="section-header">🛠 1. 통합시험</p>', unsafe_allow_html=True)
                col2.markdown('<p class="section-header">⚖️ 2. 확인검사</p>', unsafe_allow_html=True)
                col3.markdown('<p class="section-header">📊 3. 상대정확도</p>', unsafe_allow_html=True)

                for i in range(3, len(df.columns)):
                    if is_ok(target_row[i]):
                        cat = str(top_h[i])
                        name = str(sub_h[i])
                        
                        # 카드 형태의 디자인으로 출력
                        card_html = f'<div class="test-card">✅ {name}</div>'
                        
                        if "통합" in cat:
                            col1.markdown(card_html, unsafe_allow_html=True)
                        elif "확인" in cat:
                            col2.markdown(card_html, unsafe_allow_html=True)
                        elif "상대" in cat:
                            col3.markdown(card_html, unsafe_allow_html=True)
else:
    st.error("가이드북 엑셀 파일을 찾을 수 없습니다.")
