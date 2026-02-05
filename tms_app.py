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
        guide_path = next((f for f in files if '가이드북' in f), '개선내역에 따른 시험방법.xlsx')
        report_path = next((f for f in files if '1.통합시험' in f), '1.통합시험 조사표.xlsx')
        check_path = next((f for f in files if '2.확인검사' in f), '2.확인검사 조사표.xlsx')
        rel_path = next((f for f in files if '상대정확도' in f), '3.상대정확도 결과서.xlsx')
        
        guide_df = pd.read_excel(guide_path, sheet_name='★최종(가이드북)', skiprows=1)
        guide_df.iloc[:, 1] = guide_df.iloc[:, 1].ffill()
        
        report_sheets = pd.read_excel(report_path, sheet_name=None)
        check_sheets = pd.read_excel(check_path, sheet_name=None) if os.path.exists(check_path) else {}
        rel_sheets = pd.read_excel(rel_path, sheet_name=None) if os.path.exists(rel_path) else {}
        
        return guide_df, report_sheets, check_sheets, rel_sheets
    except Exception as e:
        st.error(f"⚠️ 파일을 불러올 수 없습니다: {e}")
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
                                # 시트명 매칭 로직 (공백 제거 후 비교)
                                matched_name = next((s for s in report_sheets.keys() if s.replace(" ", "") == clean_name), None)
                                if matched_name:
                                    with st.expander(f"✅ {matched_name}", expanded=False):
                                        df = report_sheets[matched_name].fillna("")
                                        st.dataframe(df, use_container_width=True)
                                        df_exp = df.copy(); df_exp.insert(0, '대분류', '통합시험'); df_exp.insert(1, '시험항목', matched_name)
                                        all_data_frames.append(df_exp)
                    else: st.info("📍 대상 아님")

                # [2단: 확인검사 - 기기교체 관련 세부 항목 적용]
                with col2:
                    st.markdown("#### 🔍 2. 확인검사")
                    check_base_names = ["외관 및 구조", "전원전압 변동", "절연저항", "공급전압의 안정성", "반복성", "제로 및 스팬 드리프트", "응답시간", "직선성", "유입전류 안정성", "간섭영향", "검출한계"]
                    
                    # '외관 및 구조' 체크 시 보여줄 수질 전용 세부 시트 목록
                    water_structure_sheets = [
                        "측정소 구조 및 설비", "시료채취조", "형식승인", 
                        "측정방법", "측정범위", "교정기능(표준물질)", "정도검사 교정일자"
                    ]
                    
                    found_check = any(is_checked(target_row.iloc[11 + i]) for i in range(len(check_base_names)))
                    if found_check:
                        st.error("📍 수행 대상")
                        for i, name in enumerate(check_base_names):
                            if is_checked(target_row.iloc[11 + i]):
                                if name == "외관 및 구조":
                                    # 외관 및 구조에 해당하는 모든 수질 세부 항목 출력
                                    for s_name in water_structure_sheets:
                                        if s_name in check_sheets:
                                            with st.expander(f"✅ {s_name}", expanded=False):
                                                df = check_sheets[s_name].fillna("")
                                                st.dataframe(df, use_container_width=True)
                                                df_exp = df.copy(); df_exp.insert(0, '대분류', '확인검사'); df_exp.insert(1, '시험항목', s_name)
                                                all_data_frames.append(df_exp)
                                        else:
                                            st.caption(f"⚠️ {s_name} 시트 없음")
                                else:
                                    # 나머지 일반 확인검사 항목 (반복성 등)
                                    if name in check_sheets:
                                        with st.expander(f"✅ {name}", expanded=False):
                                            df = check_sheets[name].fillna("")
                                            st.dataframe(df, use_container_width=True)
                                            df_exp = df.copy(); df_exp.insert(0, '대분류', '확인검사'); df_exp.insert(1, '시험항목', name)
                                            all_data_frames.append(df_exp)
                                    else:
                                        st.write(f"✅ {name}")
                    else: st.info("📍 대상 아님")

                # [3단: 상대정확도]
                with col3:
                    st.markdown("#### 📊 3. 상대정확도")
                    if is_checked(target_row.iloc[22]):
                        st.error("📍 수행 대상")
                        if rel_sheets:
                            first_sheet = list(rel_sheets.keys())[0]
                            with st.expander(f"✅ 상대정확도 시험", expanded=False):
                                df = rel_sheets[first_sheet].fillna("")
                                st.dataframe(df, use_container_width=True)
                                df_exp = df.copy(); df_exp.insert(0, '대분류', '상대정확도'); df_exp.insert(1, '시험항목', '상대정확도')
                                all_data_frames.append(df_exp)
                    else: st.info("📍 대상 아님")

                # --- 💾 엑셀 다운로드 ---
                if all_data
