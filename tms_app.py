import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="TMS 통합조사표 생성기", layout="wide")

st.title("📋 TMS 개선내역별 시험방법 일괄 확인")

@st.cache_data
def load_all_data():
    # 1. 가이드북 로드
    guide_path = '개선내역에 따른 시험방법(2025 최종).xlsx'
    guide_df = pd.read_excel(guide_path, sheet_name='★최종(가이드북)', skiprows=1)
    guide_df.iloc[:, 1] = guide_df.iloc[:, 1].ffill()
    
    # 2. 통합시험 조사표 로드
    report_path = '1.통합시험 조사표.xlsx'
    # 시트 이름을 key로, df를 value로 저장 (모든 시트 로드)
    all_sheets = pd.read_excel(report_path, sheet_name=None)
    
    # 시트 이름 매칭을 위해 공백 제거된 매핑 딕셔너리 생성
    # 예: {'1.일반현황': '1. 일반현황'}
    sheet_map = {name.replace(" ", ""): name for name in all_sheets.keys()}
    
    return guide_df, all_sheets, sheet_map

guide_df, report_sheets, sheet_map = load_all_data()

# --- 사이드바 분류 선택 ---
with st.sidebar:
    st.header("🔍 개선내역 분류")
    categories = guide_df.iloc[:, 1].dropna().unique()
    selected_cat = st.selectbox("1. 대분류 선택", categories)
    
    filtered_df = guide_df[guide_df.iloc[:, 1] == selected_cat]
    sub_items_raw = filtered_df.iloc[:, 2].dropna().unique()
    display_sub_items = [str(item).replace('\n', ' ').strip() for item in sub_items_raw]
    selected_sub_display = st.selectbox("2. 상세내역 선택", ["선택 안 함"] + display_sub_items)

def is_checked(value):
    if pd.isna(value): return False
    val_str = str(value).replace(" ", "").upper()
    return any(m in val_str for m in ['O', '○', '오', 'ㅇ'])

if selected_sub_display != "선택 안 함":
    target_row = None
    for idx, row in filtered_df.iterrows():
        if str(row.iloc[2]).replace('\n', ' ').strip() == selected_sub_display:
            target_row = row
            break

    if target_row is not None:
        st.info(f"### 📍 [{selected_sub_display}] 전체 수행 항목 결과")
        
        # 가이드북에 정의된 열 순서와 매칭될 이름들
        integrated_tests = [
            ("1. 일반현황", 3), ("2. 하드웨어 규격", 4), ("3. 소프트웨어 기능 규격", 5),
            ("4. 자료정의", 6), ("5. 측정기기 점검사항", 7), ("6. 자료생성", 8),
            ("7. 측정기기-자료수집기", 9), ("8. 자료수집기-관제센터", 10)
        ]

        col1, col2, col3 = st.columns([1.5, 0.8, 0.7], gap="medium")

        with col1:
            st.markdown("#### 💡 통합시험 상세 내용 (1~8번)")
            final_report_list = []
            
            for name, col_idx in integrated_tests:
                if is_checked(target_row.iloc[col_idx]):
                    # 🔍 시트 이름 매칭 로직 (공백 무시)
                    clean_name = name.replace(" ", "")
                    matched_sheet_name = None
                    
                    # 1. 직접 매칭 확인
                    if name in report_sheets:
                        matched_sheet_name = name
                    # 2. 공백 제거 후 매칭 확인 (예: '7.측정기기-자료수집기' vs '7. 측정기기-자료수집기')
                    elif clean_name in sheet_map:
                        matched_sheet_name = sheet_map[clean_name]
                    # 3. 앞부분 숫자만으로 매칭 확인 (예: '1.'으로 시작하는 시트)
                    else:
                        prefix = name.split('.')[0] + "."
                        for s_key in report_sheets.keys():
                            if s_key.startswith(prefix):
                                matched_sheet_name = s_key
                                break

                    if matched_sheet_name:
                        with st.expander(f"✅ {matched_sheet_name} (클릭하여 내용 보기)", expanded=True):
                            df_content = report_sheets[matched_sheet_name]
                            # 데이터 표시 (상위 15행, 결측치 제거 후 깨끗하게)
                            st.dataframe(df_content.dropna(how='all').head(15), use_container_width=True)
                            
                            temp_df = df_content.copy()
                            temp_df.insert(0, '시험구분', matched_sheet_name)
                            final_report_list.append(temp_df)
                    else:
                        st.warning(f"⚠️ '{name}' 시트를 찾을 수 없습니다. (시트명 확인 필요)")
                else:
                    st.write(f"❌ ~~{name}~~")

        # --- 확인검사 및 상대정확도는 요약만 표출 ---
        with col2:
            st.markdown("#### 🔍 확인검사")
            inspection_items = [
                ("시료채취지점", 11), ("측정소 입지조건", 12), ("측정소 구조 및 설비", 13),
                ("시료채취조", 14), ("자동시료채취기", 15), ("형식승인", 16),
                ("측정방법", 17), ("측정범위", 18), ("교정기능", 19),
                ("정도검사일자", 20), ("유량계누적값", 21)
            ]
            for n, i in inspection_items:
                if is_checked(target_row.iloc[i]):
                    st.write(f"✅ **{n}**")
                else:
                    st.caption(f"  ~~{n}~~")

        with col3:
            st.markdown("#### 📊 상대정확도")
            if is_checked(target_row.iloc[22]):
                st.error("🚨 **수행 대상**")
            else:
                st.success("✅ **대상 아님**")

        # --- 다운로드 버튼 ---
        if final_report_list:
            st.divider()
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for df in final_report_list:
                    s_name = str(df['시험구분'].iloc[0])[:30]
                    df.drop(columns=['시험구분']).to_excel(writer, index=False, sheet_name=s_name)
            
            st.download_button(
                label="📥 선택된 모든 시험 양식 다운로드",
                data=output.getvalue(),
                file_name=f"TMS_조사표_{selected_sub_display}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )