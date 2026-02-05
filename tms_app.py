import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="수질 TMS 시험 항목", layout="wide")

@st.cache_data
def load_data():
    try:
        f_list = os.listdir('.')
        g_p = next((f for f in f_list if '가이드북' in f or '시험방법' in f), None)
        if not g_p: return None
        
        # 엑셀의 헤더가 2단 구조이므로 header=None으로 읽어 처리
        df_raw = pd.read_excel(g_p, header=None)
        
        # '통합시험' 글자가 있는 행(대분류)과 그 다음 행(세부 시험명) 찾기
        top_header_idx = 0
        for i in range(len(df_raw)):
            row_str = "".join(df_raw.iloc[i].astype(str))
            if "통합시험" in row_str and "확인검사" in row_str:
                top_header_idx = i
                break
        
        # 대분류 행과 소분류(시험명) 행 추출
        top_header = df_raw.iloc[top_header_idx].ffill() # 통합시험, 확인검사...
        sub_header = df_raw.iloc[top_header_idx + 1]     # 일반현황, 점검사항...
        
        # 데이터 영역 (개선내역이 시작되는 곳)
        data_df = df_raw.iloc[top_header_idx + 2:].reset_index(drop=True)
        data_df.iloc[:, 1] = data_df.iloc[:, 1].ffill() # 분류 컬럼 병합 해제
        
        return data_df, top_header, sub_header
    except:
        return None, None, None

df, top_h, sub_h = load_data()

def is_ok(val):
    s = str(val).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

if df is not None:
    search_q = st.text_input("개선내역 입력", "")
    
    if search_q:
        # 3번째 열(index 2)에서 검색
        matches = df[df.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        
        if not matches.empty:
            matches['dp'] = matches.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            sel = st.selectbox("항목 선택", ["선택하세요"] + matches['dp'].tolist())
            
            if sel != "선택하세요":
                target_row = matches[matches['dp'] == sel].iloc[0]
                
                st.write("---")
                c1, c2, c3 = st.columns(3)
                
                # 가로 전체 열을 순회하며 'O' 체크 확인
                for col_idx in range(3, len(df.columns)):
                    cell_val = target_row[col_idx]
                    
                    if is_ok(cell_val):
                        category = str(top_h[col_idx]) # 통합시험 / 확인검사 / 상대정확도
                        test_name = str(sub_h[col_idx]) # 실제 시험명 (일반현황 등)
                        
                        if "통합시험" in category:
                            with c1:
                                if "printed" not in st.session_state: st.subheader("통합시험")
                                st.write(f"**{test_name}**")
                        elif "확인검사" in category:
                            with c2:
                                if "printed_c" not in st.session_state: st.subheader("확인검사")
                                st.write(f"**{test_name}**")
                        elif "상대정확도" in category:
                            with c3:
                                if "printed_s" not in st.session_state: st.subheader("상대정확도")
                                st.write(f"**{test_name}**")

                # 헤더가 중복되지 않게 구조적으로 배치
                with c1: st.subheader("통합시험")
                with c2: st.subheader("확인검사")
                with c3: st.subheader("상대정확도")
                
                # 실제 데이터 출력 로직 재정렬
                for i in range(3, len(df.columns)):
                    if is_ok(target_row[i]):
                        cat, name = str(top_h[i]), str(sub_h[i])
                        if "통합" in cat: c1.write(f"• {name}")
                        elif "확인" in cat: c2.write(f"• {name}")
                        elif "상대" in cat: c3.write(f"• {name}")

else:
    st.warning("가이드북 파일을 확인해주세요.")
