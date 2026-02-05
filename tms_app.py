import streamlit as st
import pandas as pd
import os
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="수질 TMS 스마트 가이드", layout="wide")

@st.cache_data
def load_all_resources():
    try:
        f_list = os.listdir('.')
        g_p = next((f for f in f_list if '가이드북' in f or '시험방법' in f), None)
        r_p = next((f for f in f_list if '1.통합' in f), None)
        c_p = next((f for f in f_list if '2.확인' in f), None)
        # '상대'라는 글자가 들어간 모든 파일을 후보로 잡음
        s_p = next((f for f in f_list if '상대' in f), None)
        
        if not g_p: return None, None, None, None
        
        df_raw = pd.read_excel(g_p, header=None)
        h_idx = 0
        for i in range(len(df_raw)):
            row_vals = [str(v) for v in df_raw.iloc[i].values]
            # '상대'라는 단어만 있어도 헤더로 인식하도록 변경
            if any("통합" in v for v in row_vals) and any("상대" in v for v in row_vals):
                h_idx = i
                break
        
        top_h = df_raw.iloc[h_idx].ffill() 
        sub_h = df_raw.iloc[h_idx + 1]     
        data_df = df_raw.iloc[h_idx + 2:].reset_index(drop=True)
        data_df.iloc[:, 1] = data_df.iloc[:, 1].ffill()
        
        r_data = pd.read_excel(r_p, sheet_name=None) if r_p else {}
        c_data = pd.read_excel(c_p, sheet_name=None) if c_p else {}
        s_data = pd.read_excel(s_p, sheet_name=None) if s_p else {}
        
        return data_df, top_h, sub_h, {"통합시험": r_data, "확인검사": c_data, "상대정확도": s_data}
    except:
        return None, None, None, None

df, top_h, sub_h, survey_data = load_all_resources()

def is_ok(val):
    s = str(val).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

st.title("🌊 수질 TMS 통합 조사표 시스템")

if df is not None:
    search_q = st.text_input("🔍 개선내역 키워드 입력", placeholder="예: 측정기기 교체")
    
    if search_q:
        matches = df[df.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        if not matches.empty:
            matches['dp'] = matches.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            sel = st.selectbox("📌 항목 선택", ["선택하세요"] + matches['dp'].tolist())
            
            if sel != "선택하세요":
                target_row = matches[matches['dp'] == sel].iloc[0]
                combined_sheets = {"통합시험": [], "확인검사": [], "상대정확도": []}
                
                st.divider()
                col1, col2, col3 = st.columns(3)
                col1.header("🛠 1. 통합시험")
                col2.header("⚖️ 2. 확인검사")
                col3.header("📊 3. 상대정확도")

                for i in range(3, len(df.columns)):
                    if is_ok(target_row[i]):
                        cat_raw = str(top_h[i])
                        name = str(sub_h[i])
                        
                        # 대분류 판별 로직 강화
                        if "통합" in cat_raw: main_cat = "통합시험"; target_col = col1
                        elif "확인" in cat_raw: main_cat = "확인검사"; target_col = col2
                        elif "상대" in cat_raw: main_cat = "상대정확도"; target_col = col3
                        else: continue

                        with target_col:
                            with st.expander(f"✅ {name}", expanded=True):
                                sheets = survey_data.get(main_cat, {})
                                found = False
                                
                                # 상대정확도의 경우, 시트명이 완벽히 일치하지 않아도 모든 시트를 검토
                                for s_name, s_df in sheets.items():
                                    # 이름이 포함되거나, 상대정확도 섹션인데 시트가 1개뿐인 경우 연결
                                    if (s_name.replace(" ","") in name.replace(" ","")) or \
                                       (name.replace(" ","") in s_name.replace(" ","")) or \
                                       (main_cat == "상대정확도"):
                                        
                                        st.dataframe(s_df.fillna(""), use_container_width=True)
                                        header_df = pd.DataFrame([[f"■ {name} ({s_name})"]], columns=[s_df.columns[0]])
                                        combined_sheets[main_cat].append(header_df)
                                        combined_sheets[main_cat].append(s_df)
                                        combined_sheets[main_cat].append(pd.DataFrame([[""]]))
                                        found = True
                                        if main_cat != "상대정확도": break # 상대정확도는 여러 시트일 수 있어 계속 진행
                                
                                if not found: st.caption("⚠️ 매칭되는 시트를 찾을 수 없습니다.")

                # 엑셀 생성
                output_xlsx = BytesIO()
                with pd.ExcelWriter(output_xlsx, engine='xlsxwriter') as writer:
                    for s_title, d_list in combined_sheets.items():
                        if d_list:
                            pd.concat(d_list, ignore_index=True).to_excel(writer, sheet_name=s_title, index=False)
                
                if any(combined_sheets.values()):
                    st.download_button("📥 통합 엑셀 다운로드", output_xlsx.getvalue(), 
                                     file_name=f"수질TMS_조사표_{sel}.xlsx")
