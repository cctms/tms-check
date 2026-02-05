import streamlit as st
import pandas as pd
from io import BytesIO
import os

st.set_page_config(page_title="TMS 수질 시험 매칭 시스템", layout="wide")

@st.cache_data
def load_all_data():
    try:
        f_list = os.listdir('.')
        # 파일 경로 탐색
        g_p = next((f for f in f_list if '가이드북' in f or '시험방법' in f), None)
        r_p = next((f for f in f_list if '1.통합' in f), None)
        c_p = next((f for f in f_list if '2.확인' in f), None)
        s_p = next((f for f in f_list if '상대' in f or '3.' in f), None)
        
        if not g_p: return None, None, None, None
        
        # 1. 가이드북 로드 (진짜 시험 항목명이 있는 행 찾기)
        # 보통 1행은 파일제목, 2행은 대분류, 3행에 실제 시험명이 있습니다.
        guide_raw = pd.read_excel(g_p, header=None)
        
        # 3행(index 2)을 실제 컬럼명으로 사용 (수질항목들이 나열된 행)
        # 만약 구조가 다를 경우를 대비해 '반복성'이나 '재현성'이 있는 행을 찾음
        header_idx = 2
        for i in range(len(guide_raw)):
            row_str = "".join(guide_raw.iloc[i].astype(str))
            if "반복성" in row_str or "제로드리프트" in row_str or "시료" in row_str:
                header_idx = i
                break
        
        df_guide = pd.read_excel(g_p, skiprows=header_idx)
        df_guide.iloc[:, 1] = df_guide.iloc[:, 1].ffill() # '분류' 채우기
        
        # 2. 조사표 파일들 로드 (탭 이름 추출용)
        r_sheets = pd.read_excel(r_p, sheet_name=None) if r_p else {}
        c_sheets = pd.read_excel(c_p, sheet_name=None) if c_p else {}
        s_sheets = pd.read_excel(s_p, sheet_name=None) if s_p else {}
        
        return df_guide, r_sheets, c_sheets, s_sheets
    except Exception as e:
        st.error(f"데이터 로드 중 오류: {e}")
        return None, None, None, None

df_guide, r_sheets, c_sheets, s_sheets = load_all_data()

def is_checked(value):
    if pd.isna(value): return False
    val = str(value).replace(" ", "").upper()
    return any(m in val for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

st.title("📋 개선내역별 시험수행항목 매칭")

if df_guide is not None:
    # 3행(개선내역) 기준으로 검색
    search_q = st.text_input("개선내역을 입력하세요 (예: 기기교체, 펌프수리 등)", "")
    
    if search_q:
        # 가이드북에서 검색 (3번째 열이 개선내역이라고 가정)
        match_rows = df_guide[df_guide.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        
        if not match_rows.empty:
            match_rows['display_name'] = match_rows.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            selected_item = st.selectbox("정확한 항목 선택", ["선택하세요"] + match_rows['display_name'].tolist())
            
            if selected_item != "선택하세요":
                target_row = match_rows[match_rows['display_name'] == selected_item].iloc[0]
                
                # 가이드북에서 'ㅇ' 표시된 모든 컬럼명 추출
                active_tests = []
                for col in df_guide.columns:
                    if is_checked(target_row[col]):
                        clean_col = str(col).strip()
                        if not any(ex in clean_col for ex in ["순번", "분류", "개선내역", "Unnamed"]):
                            active_tests.append(clean_col)
                
                st.success(f"🔍 **가이드북 기준 필요 시험:** {', '.join(active_tests)}")
                st.write("---")

                # 탭 매칭 함수
                def find_matching_sheets(check_list, sheet_dict):
                    matched = []
                    for s_name in sheet_dict.keys():
                        s_name_clean = str(s_name).replace(" ", "")
                        # 1. 가이드북 시험명이 탭 이름에 포함되는지 확인
                        if any(tc.replace(" ", "") in s_name_clean or s_name_clean in tc.replace(" ", "") for tc in check_list):
                            matched.append(s_name)
                        # 2. 예외 규칙 (예: 외관 및 구조 체크 시 관련 탭들)
                        elif "외관" in "".join(check_list) or "구조" in "".join(check_list):
                            if any(k in s_name_clean for k in ["구조", "시료", "승인", "방법", "범위", "교정", "일자"]):
                                matched.append(s_name)
                        # 3. 유량 관련
                        elif "유량" in "".join(check_list) and any(k in s_name_clean for k in ["유량", "누적"]):
                            matched.append(s_name)
                    return list(set(matched))

                c1, c2, c3 = st.columns(3)
                all_export_data = []

                with c1:
                    st.subheader("📁 통합시험")
                    matches = find_matching_sheets(active_tests, r_sheets)
                    for m in matches:
                        with st.expander(f"📑 {m}"):
                            st.dataframe(r_sheets[m].fillna(""))
                            temp_df = r_sheets[m].copy()
                            temp_df.insert(0, '탭이름', m)
                            all_export_data.append(temp_df)
                    if not matches: st.info("매칭된 탭 없음")

                with c2:
                    st.subheader("📁 확인검사")
                    matches = find_matching_sheets(active_tests, c_sheets)
                    for m in matches:
                        with st.expander(f"📑 {m}"):
                            st.dataframe(c_sheets[m].fillna(""))
                            temp_df = c_sheets[m].copy()
                            temp_df.insert(0, '탭이름', m)
                            all_export_data.append(temp_df)
                    if not matches: st.info("매칭된 탭 없음")

                with c3:
                    st.subheader("📁 상대정확도")
                    # 상대정확도는 가이드북에 해당 단어가 있을 때만 표출
                    if any("상대" in tc for tc in active_tests):
                        for m in s_sheets.keys():
                            with st.expander(f"📑 {m}"):
                                st.dataframe(s_sheets[m].fillna(""))
                                temp_df = s_sheets[m].copy()
                                temp_df.insert(0, '탭이름', m)
                                all_export_data.append(temp_df)
                    else: st.info("매칭된 탭 없음")

                if all_export_data:
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        pd.concat(all_export_data).to_excel(writer, index=False)
                    st.download_button("📥 매칭된 수행항목 다운로드", output.getvalue(), "Matched_TMS_Tasks.xlsx")
        else:
            st.warning("검색 결과가 없습니다.")
