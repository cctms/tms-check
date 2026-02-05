import streamlit as st
import pandas as pd
from io import BytesIO
import os

# 1. 페이지 설정
st.set_page_config(page_title="수질 TMS 시험항목 도구", layout="wide")

# 스타일 설정: 분석 결과 줄바꿈 방지 및 가독성 향상
st.markdown("""
    <style>
    .single-line-header {
        white-space: nowrap;
        overflow-x: auto;
        font-size: 1.6rem;
        font-weight: 700;
        padding: 10px 0px;
        color: #0E1117;
        border-bottom: 2px solid #F0F2F6;
        margin-bottom: 20px;
    }
    .stExpander {
        border: 1px solid #f0f2f6;
        margin-bottom: 5px;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("📋 수질 TMS 개선내역별 시험항목")

# 2. 데이터 로드 함수
@st.cache_data
def load_all_data():
    try:
        files = os.listdir('.')
        guide_path = next((f for f in files if '가이드북' in f or '시험방법' in f), None)
        report_path = next((f for f in files if '1.통합시험' in f), None)
        check_path = next((f for f in files if '2.확인검사' in f), None)
        rel_path = next((f for f in files if '상대정확도' in f or '3.상대정확도' in f), None)
        
        if not guide_path:
            st.error(f"❌ 가이드북 파일을 찾을 수 없습니다. 현재 폴더 파일: {files}")
            return None, None, None, None
            
        guide_df = pd.read_excel(guide_path, sheet_name='★최종(가이드북)', skiprows=1)
        guide_df.iloc[:, 1] = guide_df.iloc[:, 1].ffill()
        
        # 시트별 데이터 로드
        report_sheets = pd.read_excel(report_path, sheet_name=None) if report_path else {}
        check_sheets = pd.read_excel(check_path, sheet_name=None) if check_path else {}
        rel_sheets = pd.read_excel(rel_path, sheet_name=None) if rel_path else {}
        
        return guide_df, report_sheets, check_sheets, rel_sheets
    except Exception as e:
        st.error(f"⚠️ 데이터 로드 중 오류 발생: {e}")
        return None, None, None, None

guide_df, report_sheets, check_sheets, rel_sheets = load_all_data()

# 체크 표시 인식 함수
def is_checked(value):
    if pd.isna(value): return False
    val_str = str(value).replace(" ", "").upper()
    return any(m in val_str for m in ['O', '○', '오', 'ㅇ', 'V', 'CHECK'])

if guide_df is not None:
    st.markdown("### 🔍 개선내역 검색")
    search_query = st.text_input("키워드를 입력하세요 (예: 기기교체)", "")

    if search_query:
        # 개선내역 열(인덱스 2)에서 검색
        search_results = guide_df[guide_df.iloc[:, 2].str.contains(search_query, na=False, case=False)].copy()
        
        if not search_results.empty:
            search_results['display_name'] = search_results.apply(lambda x: f"[{x.iloc[1]}] {str(x.iloc[2]).strip()}", axis=1)
            options = search_results['display_name'].tolist()
            selected_option = st.selectbox(f"검색 결과 ({len(options)}건):", ["선택하세요"] + options)
            
            if selected_option != "선택하세요":
                target_row = search_results[search_results['display_name'] == selected_option].iloc[0]
                selected_sub = str(target_row.iloc[2]).replace('\n', ' ').strip()
                
                st.divider()
                st.markdown(f'<div class="single-line-header">🎯 분석 결과: {selected_option}</div>', unsafe_allow_html=True)
                
                all_data_frames = []
                col1, col2, col3 = st.columns([1, 1, 1])

                # [1. 통합시험]
                with col1:
                    st.markdown("#### 📝 1. 통합시험")
                    test_items = [
                        ("1. 일반현황", 3), ("2. 하드웨어 규격", 4), ("3. 소프트웨어 기능 규격", 5),
                        ("4. 자료정의", 6), ("5. 측정기기 점검사항", 7), ("6. 자료생성", 8),
                        ("7. 측정기기-자료수집기", 9), ("8. 자료수집기-관제센터", 10)
                    ]
                    
                    found_any_test = any(is_checked(target_row.iloc[idx]) for _, idx in test_items)
                    if "교체" in selected_sub: found_any_test = True # 교체 시 강제 활성

                    if found_any_test:
                        st.error("📍 수행 대상")
                        for name, col_idx in test_items:
                            # 체크되어 있거나 교체 시 7, 8번은 필수
                            if is_checked(target_row.iloc[col_idx]) or ("교체" in selected_sub and col_idx in [9, 10]):
                                # 번호 기반 시트 매칭 (예: "7."으로 시작하는 시트 찾기)
                                num_prefix = name.split('.')[0] + "."
                                matched_name = next((s for s in report_sheets.keys() if s.strip().startswith(num_prefix)), None)
                                
                                if matched_name:
                                    with st.expander(f"✅ {name}", expanded=False):
                                        df = report_sheets[matched_name].fillna("")
                                        st.dataframe(df, use_container_width=True)
                                        df_exp = df.copy()
                                        df_exp.insert(0, '대분류', '통합시험')
                                        df_exp.insert(1, '시험항목', name)
                                        all_data_frames.append(df_exp)
                                else:
                                    st.warning(f"⚠️ {name} (조사표 시트 미연결)")
                    else:
                        st.info("📍 대상 아님")

                # [2. 확인검사]
                with col2:
                    st.markdown("#### 🔍 2. 확인검사")
                    check_base_names = ["외관 및 구조", "전원전압 변동", "절연저항", "공급전압의 안정성", "반복성", "제로 및 스팬 드리프트", "응답시간", "직선성", "유입전류 안정성", "간섭영향", "검출한계"]
                    water_structure_sheets = ["측정소
