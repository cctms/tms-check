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

            all_data_frames = [] # 합칠 데이터프레임들을 담을 리스트

            st.markdown("### 📝 수행 항목")
            col_main, col_side = st.columns([2, 1])

            with col_main:
                for name, col_idx in test_items:
                    if is_checked(target_row.iloc[col_idx]):
                        clean_name = name.replace(" ", "")
                        matched_name = next((val for key, val in sheet_map.items() if key == clean_name), None) or (name if name in report_sheets else None)

                        if matched_name:
                            with st.expander(f"✅ {matched_name}", expanded=True):
                                df_content = report_sheets[matched_name].fillna("")
                                st.dataframe(df_content, use_container_width=True)
                                
                                # 구분을 위해 데이터 맨 앞에 '항목명' 컬럼을 추가해서 저장
                                temp_df = df_content.copy()
                                temp_df.insert(0, '구분', matched_name)
                                all_data_frames.append(temp_df)

            with col_side:
                st.markdown("#### 🔍 추가 확인")
                if is_checked(target_row.iloc[22]): st.error("📊 상대정확도: **대상**")
                else: st.success("📊 상대정확도: **미대상**")

            # --- 🔥 데이터 합치기 및 다운로드 로직 ---
            if all_data_frames:
                st.divider()
                
                # 모든 데이터를 하나로 합침 (세로로 이어붙이기)
                final_combined_df = pd.concat(all_data_frames, ignore_index=True)
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # 'Merged_Report'라는 단 하나의 시트에 저장
                    final_combined_df.to_excel(writer, index=False, sheet_name='TMS_통합조사표')
                    
                    # 보기 좋게 열 너비 조정
                    worksheet = writer.sheets['TMS_통합조사표']
                    worksheet.set_column(0, 0, 25) # 구분 열
                    worksheet.set_column(1, 10, 20) # 나머지 데이터 열
                
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 합쳐진 조사표 다운로드 (단일 시트)",
                    data=excel_data,
                    file_name="TMS_Combined_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
