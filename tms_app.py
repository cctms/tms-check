import streamlit as st
import pandas as pd
import os
from io import BytesIO

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
        # 상대정확도시험 확인서 파일 찾기
        s_p = next((f for f in f_list if '상대정확도' in f and '확인서' in f), None) 
        if not s_p: # 파일명이 다를 경우를 대비한 보조 검색
            s_p = next((f for f in f_list if '3.' in f or '상대정확도' in f), None)
        
        if not g_p: return None, None, None, None
        
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
        
        r_data = pd.read_excel(r_p, sheet_name=None) if r_p else {}
        c_data = pd.read_excel(c_p, sheet_name=None) if c_p else {}
        # 상대정확도시험 확인서 데이터 로드
        s_data = pd.read_excel(s_p, sheet_name=None) if s_p else {}
        
        return data_df, top_h, sub_h, {"통합시험": r_data, "확인검사": c_data, "상대정확도": s_data}
    except:
        return None, None, None, None

df, top_h, sub_h, survey_data = load_all_resources()

def is_ok(val):
    s = str(val).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

st.markdown('<p class="main-title">🌊 수질 TMS 맞춤 통합 조사표 생성</p>', unsafe_allow_html=True)

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
                combined_sheets = {"통합시험": [], "확인검사": [], "상대정확도": []}
                
                st.write("---")
                col1, col2, col3 = st.columns(3)
                col1.markdown('<p class="section-header">🛠 1. 통합시험</p>', unsafe_allow_html=True)
                col2.markdown('<p class="section-header">⚖️ 2. 확인검사</p>', unsafe_allow_html=True)
                col3.markdown('<p class="section-header">📊 3. 상대정확도</p>', unsafe_allow_html=True)

                for i in range(3, len(df.columns)):
                    if is_ok(target_row[i]):
                        cat, name = str(top_h[i]), str(sub_h[i])
                        main_cat = "통합시험" if "통합" in cat else "확인검사" if "확인" in cat else "상대정확도" if "상대" in cat else None
                        if not main_cat: continue
                        
                        target_col = col1 if "통합시험" == main_cat else col2 if "확인검사" == main_cat else col3
                        
                        with target_col:
                            with st.expander(f"✅ {name}"):
                                sheets = survey_data.get(main_cat, {})
                                found = False
                                for s_name, s_df in sheets.items():
                                    # 상대정확도의 경우 '확인서' 관련 시트를 우선 매칭
                                    if s_name.replace(" ","") in name.replace(" ","") or name.replace(" ","") in s_name.replace(" ",""):
                                        st.dataframe(s_df.fillna(""), use_container_width=True)
                                        header_df = pd.DataFrame([[f"■ {name}"]], columns=[s_df.columns[0] if not s_df.empty else "항목"])
                                        combined_sheets[main_cat].append(header_df)
                                        combined_sheets[main_cat].append(s_df)
                                        combined_sheets[main_cat].append(pd.DataFrame([[""]]))
                                        found = True
                                        break
                                if not found: st.info("데이터 없음")

                # --- 통합 엑셀 파일 생성 ---
                output_xlsx = BytesIO()
                with pd.ExcelWriter(output_xlsx, engine='xlsxwriter') as writer:
                    for sheet_name, df_list in combined_sheets.items():
                        if df_list:
                            final_df = pd.concat(df_list, ignore_index=True)
                            final_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                st.divider()
                if any(combined_sheets.values()):
                    st.download_button(
                        label=f"📥 {sel} 통합 조사표 다운로드",
                        data=output_xlsx.getvalue(),
                        file_name=f"수질TMS_통합조사표_{sel.replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
else:
    st.error("파일 로드 실패")
