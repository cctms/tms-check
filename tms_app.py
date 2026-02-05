import streamlit as st
import pandas as pd
import os
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="수질 TMS 스마트 가이드", layout="wide")

# 디자인 CSS: 제목 크기를 3.0rem으로 키우고 스타일을 강화했습니다.
st.markdown("""
    <style>
    .main-title { 
        font-size: 3.0rem; 
        font-weight: 900; 
        color: #1E3A8A; 
        text-align: center; 
        margin-top: 1rem;
        margin-bottom: 3rem; 
        text-shadow: 1px 1px 2px #d1d1d1;
    }
    .section-header { 
        background: #1E3A8A; 
        color: white; 
        padding: 12px; 
        border-radius: 8px; 
        text-align: center; 
        font-weight: 700; 
        font-size: 1.2rem;
        margin-bottom: 15px; 
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

# 수정된 메인 제목
st.markdown('<p class="main-title">🌊 수질TMS 개선내역에 따른 통합 조사표</p>', unsafe_allow_html=True)

if df is not None:
    c_left, c_mid, c_right = st.columns([1, 2, 1])
    with c_mid:
        search_q = st.text_input("🔍 개선내역 키워드 입력 (예: 측정기기 교체)", placeholder="검색어를 입력하세요")
    
    if search_q:
        matches = df[df.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        if not matches.empty:
            matches['dp'] = matches.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            with c_mid:
                sel = st.selectbox("📌 해당되는 개선내역 항목을 선택하세요", ["선택하세요"] + matches['dp'].tolist())
            
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
                        
                        # 상대정확도 nan 처리 및 섹션 분류
                        if "상대" in cat_raw:
                            main_cat = "상대정확도"
                            target_col = col3
                            if name.lower() in ['nan', '', 'none']:
                                name = "상대정확도시험"
                        elif "통합" in cat_raw:
                            main_cat = "통합시험"
                            target_col = col1
                        elif "확인" in cat_raw:
                            main_cat = "확인검사"
                            target_col = col2
                        else:
                            continue

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
                                if not found: st.caption("⚠️ 시트 매칭 실패")

                # 통합 엑셀 생성
                output_xlsx = BytesIO()
                with pd.ExcelWriter(output_xlsx, engine='xlsxwriter') as writer:
                    for s_title, d_list in combined_sheets.items():
                        if d_list:
                            pd.concat(d_list, ignore_index=True).to_excel(writer, sheet_name=s_title, index=False)
                
                st.divider()
                if any(combined_sheets.values()):
                    st.download_button(
                        label=f"📥 {sel} 관련 통합 조사표 다운로드",
                        data=output_xlsx.getvalue(),
                        file_name=f"수질TMS_통합조사표_{sel.replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
else:
    st.error("데이터 파일을 찾을 수 없습니다.")
