import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="수질 TMS 시험 항목", layout="wide")

@st.cache_data
def load_data():
    try:
        f_list = os.listdir('.')
        g_p = next((f for f in f_list if '가이드북' in f or '시험방법' in f), None)
        if not g_p: return None, None, None
        
        # 가이드북 로드 (2단 헤더 구조 처리)
        df_raw = pd.read_excel(g_p, header=None)
        
        # '통합시험'과 '확인검사'가 동시에 존재하는 행을 대분류 행으로 인식
        h_idx = 0
        for i in range(len(df_raw)):
            row_vals = df_raw.iloc[i].astype(str).values
            if any("통합시험" in v for v in row_vals) and any("확인검사" in v for v in row_vals):
                h_idx = i
                break
        
        # 1행: 통합시험, 확인검사 등 (대분류)
        # 2행: 일반현황, 점검사항 등 (시험명)
        top_h = df_raw.iloc[h_idx].ffill() 
        sub_h = df_raw.iloc[h_idx + 1]     
        
        # 데이터 영역 (실제 개선내역 데이터)
        data_df = df_raw.iloc[h_idx + 2:].reset_index(drop=True)
        data_df.iloc[:, 1] = data_df.iloc[:, 1].ffill() # 분류 병합 해제
        
        return data_df, top_h, sub_h
    except:
        return None, None, None

df, top_h, sub_h = load_data()

def is_ok(val):
    s = str(val).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

st.title("📋 개선내역별 수행 시험 항목")

if df is not None:
    search_q = st.text_input("개선내역 입력 (예: 측정기기 교체)", "")
    
    if search_q:
        # 3번째 열(개선내역)에서 검색
        matches = df[df.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        
        if not matches.empty:
            matches['dp'] = matches.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            sel = st.selectbox("항목 선택", ["선택하세요"] + matches['dp'].tolist())
            
            if sel != "선택하세요":
                target_row = matches[matches['dp'] == sel].iloc[0]
                
                st.write("---")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.subheader("1. 통합시험")
                with col2:
                    st.subheader("2. 확인검사")
                with col3:
                    st.subheader("3. 상대정확도")

                # 가로 열을 순차적으로 돌며 'O'가 있으면 해당 섹션에 시험명 기입
                for i in range(3, len(df.columns)):
                    if is_ok(target_row[i]):
                        cat = str(top_h[i])   # 대분류 (통합/확인/상대)
                        name = str(sub_h[i])  # 시험명 (일반현황 등)
                        
                        if "통합" in cat:
                            col1.write(f"• **{name}**")
                        elif "확인" in cat:
                            col2.write(f"• **{name}**")
                        elif "상대" in cat:
                            col3.write(f"• **{name}**")
else:
    st.error("가이드북 파일을 찾을 수 없습니다. 파일명을 확인해주세요.")
