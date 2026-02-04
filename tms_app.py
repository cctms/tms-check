import streamlit as st
import pandas as pd
from io import BytesIO

# 1. 페이지 설정
st.set_page_config(page_title="TMS 통합조사표 생성기", layout="wide")

st.title("📋 TMS 개선내역별 시험방법 발췌 도구")

# 2. 데이터 로드 함수
@st.cache_data
def load_all_data():
    try:
        # 파일 경로 (GitHub 저장소 내 파일명과 일치해야 함)
        guide_path = '개선내역에 따른 시험방법(2025 최종).xlsx'
        report_path = '1.통합시험 조사표.xlsx'
        
        guide_df = pd.read_excel(guide_path, sheet_name='★최종(가이드북)', skiprows=1)
        guide_df.iloc[:, 1] = guide_df.iloc[:, 1].ffill()
        
        report_sheets = pd.read_excel(report_path, sheet_name=None)
        sheet_map = {name.replace(" ", ""): name for name in report_sheets.keys()}
        
        return guide_df, report_sheets, sheet_map
    except Exception as e:
        st.error(f"⚠️ 파일을 불러올 수 없습니다: {e}")
        return None, None, None

guide_df, report_sheets, sheet_map = load_all_data()

def is_checked(value):
    if pd.isna(value): return False
    val_str = str(value).replace(" ", "").upper()
    return any(m in val_str for m in ['O', '○', '오', 'ㅇ', 'V'])

if guide_df is not None:
    with st.sidebar:
        st.header("🔍 개선내역 선택")
        categories = guide_df.iloc[:, 1].dropna().unique()
        selected_cat = st.selectbox("1. 대분류", categories)
        
        filtered_df = guide_df[guide_df.iloc[:, 1] == selected_cat]
        sub_items = [str(item).replace('\n', ' ').strip() for item in filtered_df.iloc[:, 2].dropna().unique()]
        selected_sub = st.selectbox("2. 상세내역", ["선택 안 함"] + sub_items)

    if selected_sub != "선택 안 함":
        target_row = next((row for _, row in filtered_df.iterrows() if str(row.iloc[2]).replace('\n', ' ').strip() == selected_sub), None)

        if target_row is not None:
            st.success(f"🎯 **선택:** {selected_sub}")
            
            test_items = [
                ("1. 일반현황", 3), ("2. 하드웨어 규격", 4), ("3. 소프트웨어 기능 규격", 5),
                ("4. 자료정의", 6), ("5. 측정기기 점검사항", 7), ("6. 자료생성", 8),
                ("7. 측정기기-자료수집기", 9), ("8. 자료수집기-관제센터", 10)
            ]

            final_dfs = [] 

            st.markdown("### 📝 수행 항목")
            col_main, col_side = st.columns([2, 1])

            with col_main:
                for name, col_idx in test_items:
                    if is_checked(target_row.iloc[col_idx]):
                        clean_name = name.replace(" ", "")
                        matched_name = next((val for key, val in sheet_map.items() if key == clean_name), None) or (name if name in report_sheets else None)

                        if matched_name:
                            with st.expander(f"✅ {matched_name}", expanded=True):
                                # 뷰어 호환성을 위해 결측값을 빈칸으로 처리
                                df_content = report_sheets[matched_name].fillna("")
                                st.dataframe(df_content, use_container_width=True)
                                final_dfs.append((matched_name, df_content))

            with col_side:
                st.markdown("#### 🔍 추가 확인")
                if is_checked(target_row.iloc[22]): st.error("📊 상대정확도: **대상**")
                else: st.success("📊 상대정확도: **미대상**")

            # --- 🛠️ 엑셀 뷰어 호환 다운로드 로직 ---
            if final_dfs:
                st.divider()
                output = BytesIO()
                
                # 호환성이 가장 높은 xlsxwriter 엔진 사용
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for s_name, df in final_dfs:
                        # 시트 이름에서 특수문자 제거 및 길이 제한 (뷰어 에러 방지)
                        safe_name = "".join([c for c in s_name if c.isalnum() or c in ' ._-'])[:31]
                        df.to_excel(writer, index=False, sheet_name=safe_name)
                        
                        # 열 너비 자동 조정 (뷰어에서 보기 편하게)
                        worksheet = writer.sheets[safe_name]
                        for i, col in enumerate(df.columns):
                            worksheet.set_column(i, i, 20)
                
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 엑셀 뷰어용 파일 다운로드",
                    data=excel_data,
                    file_name="TMS_REPORT.xlsx", # 호환성을 위해 영어 파일명 권장
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
