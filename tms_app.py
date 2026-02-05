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
    .stDownloadButton { text-align: center; }
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
        
        # 상세 조사표 데이터 (모든 시트)
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

st.markdown('<p class="main-title">🌊 수질 TMS 수행항목 & 맞춤 조사표 생성</p>', unsafe_allow_html=True)

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
                
                # --- 다운로드 로직 시작 ---
                output_xlsx = BytesIO()
                with pd.ExcelWriter(output_xlsx, engine='xlsxwriter') as writer:
                    selected_sheets_count = 0
                    
                    st.write("---")
                    col1, col2, col3 = st.columns(3)
                    col1.markdown('<p class="section-header">🛠 1. 통합시험</p>', unsafe_allow_html=True)
                    col2.markdown('<p class="section-header">⚖️ 2. 확인검사</p>', unsafe_allow_html=True)
                    col3.markdown('<p class="section-header">📊 3. 상대정확도</p>', unsafe_allow_html=True)

                    for i in range(3, len(df.columns)):
                        if is_ok(target_row[i]):
                            cat, name = str(top_h[i]), str(sub_h[i])
                            target_col = col1 if "통합" in cat else col2 if "확인" in cat else col3
                            group_key = "통합" if "통합" in cat else "확인" if "확인" in cat else "상대"
                            
                            with target_col:
                                with st.expander(f"✅ {name}"):
                                    sheets = survey_data.get(group_key, {})
                                    found = False
                                    for s_name, s_df in sheets.items():
                                        # 이름 매칭
                                        if s_name.replace(" ","") in name.replace(" ","") or name.replace(" ","") in s_name.replace(" ",""):
                                            st.dataframe(s_df.fillna(""), use_container_width=True)
                                            # 엑셀 파일에 시트 추가 (시트명 제약 31자 처리)
                                            safe_sheet_name = "".join(x for x in s_name if x.isalnum() or x in " -_")[:30]
                                            s_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                                            selected_sheets_count += 1
                                            found = True
                                            break
                                    if not found:
                                        st.info("시트 없음")
                
                # 파일 생성 후 다운로드 버튼 배치
                st.divider()
                if selected_sheets_count > 0:
                    st.download_button(
                        label=f"📥 {sel} 관련 조사표 엑셀 다운로드 ({selected_sheets_count}개 시트)",
                        data=output_xlsx.getvalue(),
                        file_name=f"수질TMS_조사표_{sel.replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                # --- 다운로드 로직 끝 ---
else:
    st.error("파일 로드 실패")
