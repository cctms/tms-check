import streamlit as st
import pandas as pd
import os
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="수질 TMS 스마트 가이드", layout="wide")

# 디자인 CSS
st.markdown("""
    <style>
    .super-title { 
        font-size: 56px !important; 
        font-weight: 800 !important; 
        color: #1E3A8A !important; 
        text-align: center !important; 
        margin-top: 30px !important;
        margin-bottom: 20px !important; 
    }
    .chat-sub { text-align: center; color: #666; font-size: 1.2rem; margin-bottom: 40px; }
    .section-header { 
        background: #1E3A8A; color: white; padding: 12px; border-radius: 8px; 
        text-align: center; font-weight: 700; font-size: 20px; margin-bottom: 15px; 
    }
    .stTextInput > div > div > input {
        border-radius: 25px !important; padding: 15px 25px !important; border: 2px solid #1E3A8A !important;
    }
    /* 다운로드 버튼 강조 스타일 */
    .stDownloadButton > button {
        width: 100% !important;
        background-color: #28a745 !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 10px !important;
        height: 3em !important;
    }
    </style>
    """, unsafe_allow_html=True)

@st.cache_data
def load_all_resources():
    try:
        f_list = os.listdir('.')
        g_p = next((f for f in f_list if '가이드북' in f or '시험방법' in f), None)
        r_p = next((f for f in f_list if '1.통합' in f), None)
        c_p = next((f for f in f_list if '2.확인' in f), None)
        s_p = next((f for f in f_list if '상대' in f), None)
        
        if not g_p: return None, None, None, None
        
        df_raw = pd.read_excel(g_p, header=None)
        h_idx = 0
        for i in range(len(df_raw)):
            row_vals = [str(v) for v in df_raw.iloc[i].values]
            if any("통합" in v for v in row_vals) and any("확인" in v for v in row_vals):
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
    except: return None, None, None, None

df, top_h, sub_h, survey_data = load_all_resources()

def is_ok(val):
    s = str(val).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

st.markdown('<div class="super-title">수질TMS 개선내역에 따른 통합 조사표</div>', unsafe_allow_html=True)
st.markdown('<p class="chat-sub">안녕하세요! 무엇을 도와드릴까요? (예: 측정기기 교체, 펌프 수리 등)</p>', unsafe_allow_html=True)

if df is not None:
    c_left, c_mid, c_right = st.columns([1, 2, 1])
    with c_mid:
        user_input = st.text_input("💬 질문하기", placeholder="발생한 개선사항을 편하게 적어주세요.")
    
    if user_input:
        search_words = ["교체", "수리", "이전", "신규", "부품", "오버홀", "전송", "변경"]
        found_keywords = [w for w in search_words if w in user_input]
        if not found_keywords:
            found_keywords = [k for k in user_input.split() if len(k) > 1]

        mask = pd.Series([False] * len(df))
        for kw in found_keywords:
            mask |= df.iloc[:, 2].astype(str).str.contains(kw, na=False)
            
        matches = df[mask]
        
        if not matches.empty:
            matches['dp'] = matches.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            with c_mid:
                st.info(f"🧐 '{', '.join(found_keywords)}' 관련 내용을 찾았습니다.")
                sel = st.selectbox("가장 적절한 항목을 선택해 주세요:", ["선택하세요"] + matches['dp'].tolist())
            
            if sel != "선택하세요":
                target_row = matches[matches['dp'] == sel].iloc[0]
                combined_sheets = {"통합시험": [], "확인검사": [], "상대정확도": []}
                
                # 데이터 미리 수집 (다운로드 버튼을 먼저 띄우기 위해)
                for i in range(3, len(df.columns)):
                    if is_ok(target_row[i]):
                        cat_raw = str(top_h[i]); name = str(sub_h[i])
                        if "상대" in cat_raw: m_cat = "상대정확도"
                        elif "통합" in cat_raw: m_cat = "통합시험"
                        elif "확인" in cat_raw: m_cat = "확인검사"
                        else: continue
                        
                        sheets = survey_data.get(m_cat, {})
                        for s_name, s_df in sheets.items():
                            if (m_cat == "상대정확도") or (s_name.replace(" ","") in name.replace(" ","")) or (name.replace(" ","") in s_name.replace(" ","")):
                                header_df = pd.DataFrame([[f"■ {name}"]], columns=[s_df.columns[0] if not s_df.empty else "항목"])
                                combined_sheets[m_cat].extend([header_df, s_df, pd.DataFrame([[""]])])
                                if m_cat != "상대정확도": break

                # --- [상단 다운로드 영역] ---
                output_xlsx = BytesIO()
                with pd.ExcelWriter(output_xlsx, engine='xlsxwriter') as writer:
                    for s_title, d_list in combined_sheets.items():
                        if d_list:
                            pd.concat(d_list, ignore_index=True).to_excel(writer, sheet_name=s_title, index=False)
                
                with c_mid:
                    st.success(f"✅ 조사가 완료되었습니다. 바로 아래 버튼으로 파일을 받으세요!")
                    st.download_button(
                        label="📥 클릭하여 통합 조사표(Excel) 다운로드",
                        data=output_xlsx.getvalue(),
                        file_name=f"수질TMS_통합조사표_{sel.replace(' ', '_')}.xlsx",
                        key="top_download_btn"
                    )
                # --------------------------

                st.write("---")
                st.caption("💡 아래는 조사표의 상세 미리보기입니다.")
                col1, col2, col3 = st.columns(3)
                col1.markdown('<p class="section-header">🛠 1. 통합시험</p>', unsafe_allow_html=True)
                col2.markdown('<p class="section-header">⚖️ 2. 확인검사</p>', unsafe_allow_html=True)
                col3.markdown('<p class="section-header">📊 3. 상대정확도</p>', unsafe_allow_html=True)

                # 하단 미리보기 영역 재출력
                for i in range(3, len(df.columns)):
                    if is_ok(target_row[i]):
                        cat_raw = str(top_h[i]); name = str(sub_h[i])
                        if "상대" in cat_raw: m_cat = "상대정확도"; t_col = col3
                        elif "통합" in cat_raw: m_cat = "통합시험"; t_col = col1
                        elif "확인" in cat_raw: m_cat = "확인검사"; t_col = col2
                        else: continue

                        with t_col:
                            with st.expander(f"✅ {name}", expanded=False):
                                sheets = survey_data.get(m_cat, {})
                                for s_name, s_df in sheets.items():
                                    if (m_cat == "상대정확도") or (s_name.replace(" ","") in name.replace(" ","")) or (name.replace(" ","") in s_name.replace(" ","")):
                                        st.dataframe(s_df.fillna(""), use_container_width=True)
                                        if m_cat != "상대정확도": break
        else:
            with c_mid:
                st.warning("단어를 조금만 더 단순하게 입력해 보시겠어요? (예: 기기 교체, 펌프 수리 등)")
