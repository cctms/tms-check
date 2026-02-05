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
        line-height: 1.2 !important;
        display: block !important;
    }
    .chat-sub {
        text-align: center;
        color: #666;
        font-size: 1.2rem;
        margin-bottom: 40px;
    }
    .section-header { 
        background: #1E3A8A; 
        color: white; 
        padding: 12px; 
        border-radius: 8px; 
        text-align: center; 
        font-weight: 700; 
        font-size: 20px;
        margin-bottom: 15px; 
    }
    /* 챗봇 느낌의 입력창 스타일 */
    .stTextInput > div > div > input {
        border-radius: 25px !important;
        padding: 15px 25px !important;
        border: 2px solid #1E3A8A !important;
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
    except:
        return None, None, None, None

df, top_h, sub_h, survey_data = load_all_resources()

def is_ok(val):
    s = str(val).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

# 제목 및 챗봇 인사말
st.markdown('<div class="super-title">수질TMS 개선내역에 따른 통합 조사표</div>', unsafe_allow_html=True)
st.markdown('<p class="chat-sub">안녕하세요! 어떤 개선사항이 발생했나요? 아래에 질문해 주세요. 👋</p>', unsafe_allow_html=True)

if df is not None:
    c_left, c_mid, c_right = st.columns([1, 2, 1])
    with c_mid:
        # 질문형 인터페이스
        user_input = st.text_input("💬 질문하기", placeholder="예: 측정기기를 교체했는데 어떤 시험을 해야 하나요?")
    
    if user_input:
        # 간단한 형태소 분석 대용 (공백 기준 핵심 키워드 검색)
        keywords = [k for k in user_input.split() if len(k) > 1]
        
        # 키워드 중 하나라도 포함된 항목 찾기
        mask = pd.Series([False] * len(df))
        for kw in keywords:
            mask |= df.iloc[:, 2].astype(str).str.contains(kw, na=False)
            
        matches = df[mask]
        
        if not matches.empty:
            matches['dp'] = matches.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            with c_mid:
                st.info(f"🧐 질문하신 내용과 관련된 {len(matches)}개의 개선내역을 찾았습니다.")
                sel = st.selectbox("가장 적절한 항목을 선택해 주세요:", ["선택하세요"] + matches['dp'].tolist())
            
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
                        cat_raw = str(top_h[i])
                        name = str(sub_h[i])
                        
                        if "상대" in cat_raw:
                            main_cat = "상대정확도"; target_col = col3
                            if name.lower() in ['nan', '', 'none']: name = "상대정확도시험"
                        elif "통합" in cat_raw:
                            main_cat = "통합시험"; target_col = col1
                        elif "확인" in cat_raw:
                            main_cat = "확인검사"; target_col = col2
                        else: continue

                        with target_col:
                            with st.expander(f"✅ {name}", expanded=False):
                                sheets = survey_data.get(main_cat, {})
                                found = False
                                for s_name, s_df in sheets.items():
                                    if (main_cat == "상대정확도") or (s_name.replace(" ","") in name.replace(" ","")) or (name.replace(" ","") in s_name.replace(" ","")):
                                        st.dataframe(s_df.fillna(""), use_container_width=True)
                                        header_df = pd.DataFrame([[f"■ {name}"]], columns=[s_df.columns[0] if not s_df.empty else "항목"])
                                        combined_sheets[main_cat].append(header_df)
                                        combined_sheets[main_cat].append(s_df)
                                        combined_sheets[main_cat].append(pd.DataFrame([[""]]))
                                        found = True
                                        if main_cat != "상대정확도": break
                                if not found: st.caption("⚠️ 데이터 매칭 실패")

                output_xlsx = BytesIO()
                with pd.ExcelWriter(output_xlsx, engine='xlsxwriter') as writer:
                    for s_title, d_list in combined_sheets.items():
                        if d_list:
                            pd.concat(d_list, ignore_index=True).to_excel(writer, sheet_name=s_title, index=False)
                
                st.divider()
                if any(combined_sheets.values()):
                    st.success(f"✅ 선택하신 '{sel}' 항목에 대한 통합 조사표가 준비되었습니다.")
                    st.download_button(
                        label="📥 통합 조사표 다운로드",
                        data=output_xlsx.getvalue(),
                        file_name=f"수질TMS_통합조사표_{sel.replace(' ', '_')}.xlsx"
                    )
        else:
            with c_mid:
                st.warning("죄송합니다. 질문하신 내용과 관련된 개선내역을 찾지 못했어요. 핵심 단어(예: 교체, 수리, 이전)를 포함해 다시 말씀해 주시겠어요?")
else:
    st.error("데이터 파일을 불러올 수 없습니다.")
