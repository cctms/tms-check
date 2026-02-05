import streamlit as st
import pandas as pd
from io import BytesIO
import os

# 1. 페이지 설정
st.set_page_config(page_title="수질 TMS 시험항목 도구", layout="wide")
st.title("📋 수질 TMS 개선내역별 시험항목")

# 2. 데이터 로드 함수
@st.cache_data
def load_all_data():
    try:
        files = os.listdir('.')
        # 파일명을 유연하게 찾기 (가이드북, 통합시험, 확인검사, 상대정확도 키워드 기준)
        guide_path = next((f for f in files if '가이드북' in f), '개선내역에 따른 시험방법.xlsx')
        report_path = next((f for f in files if '1.통합시험' in f), '1.통합시험 조사표.xlsx')
        check_path = next((f for f in files if '2.확인검사' in f), '2.확인검사 조사표.xlsx')
        rel_path = next((f for f in files if '상대정확도' in f), '3.상대정확도 결과서.xlsx')
        
        # 가이드북 로드
        guide_df = pd.read_excel(guide_path, sheet_name='★최종(가이드북)', skiprows=1)
        guide_df.iloc[:, 1] = guide_df.iloc[:, 1].ffill()
        
        # 각 조사표 시트 로드
        report_sheets = pd.read_excel(report_path, sheet_name=None)
        check_sheets = pd.read_excel(check_path, sheet_name=None) if os.path.exists(check_path) else {}
        rel_sheets = pd.read_excel(rel_path, sheet_name=None) if os.path.exists(rel_path) else {}
        
        return guide_df, report_sheets, check_sheets, rel_sheets
    except Exception as e:
        st.error(f"⚠️ 파일을 불러올 수 없습니다. 파일명을 확인해주세요: {e}")
        return None, None, None, None

guide_df, report_sheets, check_sheets, rel_sheets = load_all_data()

def is_checked(value):
    if pd.isna(value): return False
    val_str = str(value).replace(" ", "").upper()
    return any(m in val_str for m in ['O', '○', '오', 'ㅇ', 'V'])

if guide_df is not None:
    st.markdown("### 🔍 개선내역 검색")
    search_query = st.text_input("찾으시는 개선내역(예: 기기교체)을 입력하세요", "")

    if search_query:
        search_results = guide_df[guide_df.iloc[:, 2].str.contains(search_query, na=False, case=False)].copy()
        
        if not search_results.empty:
            search_results['display_name'] = search_results.apply(lambda x: f"[{x.iloc[1]}] {str(x.iloc[2]).strip()}", axis=1)
            options = search_results['display_name'].tolist()
            selected_option = st.selectbox(f"검색 결과 ({len(options)}건):", ["선택하세요"] + options)
            
            if selected_option != "선택하세요":
                target_row = search_results[search_results['display_name'] == selected_option].iloc[0]
                selected_sub = str(target_row.iloc[2]).replace('\n', ' ').strip()
                
                st.divider()
                st.markdown(f"### 🎯 분석 결과: {selected_option}")
                
                all_data_frames = []
                col1, col2, col3 = st.columns([1, 1, 1])

                # [1단: 통합시험]
                with col1:
                    st.markdown("#### 📝 1. 통합시험")
                    test_items = [("1. 일반현황", 3), ("2. 하드웨어 규격", 4), ("3. 소프트웨어 기능 규격", 5),
                                  ("4. 자료정의", 6), ("5. 측정기기 점검사항", 7), ("6. 자료생성", 8),
                                  ("7. 측정기기-자료수집기", 9), ("8. 자료수집기-관제센터", 10)]
                    found_test = any(is_checked(target_row.iloc[col_idx]) for _, col_idx in test_items)
                    if found_test:
                        st.error("📍 수행 대상")
                        for name, col_idx in test_items:
                            if is_checked(target_row.iloc[col_idx]):
                                clean_name = name.replace(" ", "")
                                matched_name = next((s for s in report_sheets.keys() if s.replace(" ", "") == clean_name), None)
                                if matched_name:
                                    with st.expander(f"✅ {matched_name}", expanded=False):
                                        df = report_sheets[matched_name].fillna("")
                                        st.dataframe(df, use_container_width=True)
                                        df_exp = df.copy(); df_exp.insert(0, '대분류', '통합시험');
