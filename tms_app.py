import streamlit as st
import pandas as pd
from io import BytesIO

# 1. 페이지 설정
st.set_page_config(page_title="TMS 통합조사표 생성기", layout="wide")

st.title("📋 TMS 개선내역별 시험방법 일괄 확인")

# 2. 데이터 로드 함수 (캐싱 적용으로 속도 향상)
@st.cache_data
def load_all_data():
    try:
        # 가이드북 로드 (파일명 확인 필수)
        guide_path = '개선내역에 따른 시험방법(2025 최종).xlsx'
        guide_df = pd.read_excel(guide_path, sheet_name='★최종(가이드북)', skiprows=1)
        # 대분류 빈칸 채우기
        guide_df.iloc[:, 1] = guide_df.iloc[:, 1].ffill()
        
        # 통합시험 조사표 로드 (파일명 확인 필수)
        report_path = '1.통합시험 조사표.xlsx'
        # 모든 시트를 딕셔너리 형태로 로드 {시트명: 데이터프레임}
        all_sheets = pd.read_excel(report_path, sheet_name=None)
        
        # 시트 매칭을 위해 공백 제거된 매핑 생성
        sheet_map = {name.replace(" ", ""): name for name in all_sheets.keys()}
        
        return guide_df, all_sheets, sheet_map
    except Exception as e:
        st.error(f"⚠️ 파일 로드 중 오류 발생: {e}. GitHub에 파일이 있는지 확인하세요.")
        return None, None, None

guide_df, report_sheets, sheet_map = load_all_data()

# 데이터 로드 성공 시에만 실행
if guide_df is not None:
    # --- 사이드바 분류 선택 ---
    with st.sidebar:
        st.header("🔍 개선내역 분류")
        categories = guide_df.iloc[:, 1].dropna().unique()
        selected_cat = st.selectbox("1. 대분류 선택", categories)
        
        filtered_df = guide_df[guide_df.iloc[:, 1] == selected_cat]
        sub_items_raw = filtered_df.iloc[:, 2].dropna().unique()
        display_sub_items = [str(item).replace('\n', ' ').strip() for item in sub_items_raw]
        selected_sub_display = st.selectbox("2. 상세내역 선택", ["선택 안 함"] + display_sub_items)

    # 체크 표시 확인 함수 (O, ○, ㅇ 등 대응)
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
            
            # 1~8번 통합시험 항목 및 가이드북 열 인덱스
            integrated_tests = [
                ("1. 일반현황", 3), ("2. 하드웨어 규격", 4), ("3. 소프트웨어 기능 규격", 5),
                ("4. 자료정의", 6), ("5. 측정기기 점검사항", 7), ("6. 자료생성", 8),
                ("7. 측정기기-자료수집기", 9), ("8. 자료수집기-관제센터", 10)
            ]

            col1, col2, col3 = st.columns([1.5, 0.8, 0.7], gap="medium")

            # 엑셀 생성을 위한 리스트
            final_report_list = []

            with col1:
                st.markdown("#### 💡 통합시험 상세 내용 (1~8번)")
                
                for name, col_idx in integrated_tests:
                    if is_checked(target_row.iloc[col_idx]):
                        clean_name = name.replace(" ", "")
                        matched_sheet_name = None
                        
                        # 시트 매칭 로직
                        if name in report_sheets:
                            matched_sheet_name = name
                        elif clean_name in sheet_map:
                            matched_sheet_name = sheet_map[clean_name]
                        else:
                            prefix = name.split('.')[0] + "."
                            for s_key in report_sheets.keys():
                                if s_key.startswith(prefix):
                                    matched_sheet_name = s_key
                                    break

                        if matched_sheet_name:
                            with st.expander(f"✅ {matched_sheet_name} (내용 보기)", expanded=True):
                                df_content = report_sheets[matched_sheet_name]
                                # 화면용 미리보기 (상위 15행)
                                st.dataframe(df_content.dropna(how='all').head(15), use_container_width=True)
                                
                                # 엑셀 병합용 데이터 저장
                                temp_df = df_content.copy()
                                temp_df._sheet_name = matched_sheet_name # 시트 이름 임시 저장
                                final_report_list.append(temp_df)
                        else:
                            st.warning(f"⚠️ '{name}' 시트를 찾을 수 없습니다.")
                    else:
                        st.write(f"❌ ~~{name}~~")

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

            # --- 엑셀 병합 및 다운로드 버튼 ---
            if final_report_list:
                st.divider()
                output = BytesIO()
                # xlsxwriter를 사용하여 여러 시트로 저장
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for df in final_report_list:
                        s_name = df._sheet_name[:31] # 시트명 글자수 제한 대응
                        df.to_excel(writer, index=False, sheet_name=s_name)
                
                st.download_button(
                    label=f"📥 {selected_sub_display} 조사표 엑셀 다운로드",
                    data=output.getvalue(),
                    file_name=f"TMS_조사표_{selected_sub_display}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.info("왼쪽 사이드바에서 개선내역 상세항목을 선택해 주세요.")
