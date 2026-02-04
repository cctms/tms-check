import streamlit as st
import pandas as pd
from io import BytesIO

# 1. 페이지 설정
st.set_page_config(page_title="TMS 시험항목 도구", layout="wide")

# 요청하신 대로 타이틀 변경
st.title("📋 TMS 개선내역별 시험항목")

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
            
            all_data_frames = []

            # --- 1. 통합시험 항목 ---
            test_items = [
                ("1. 일반현황", 3), ("2. 하드웨어 규격", 4), ("3. 소프트웨어 기능 규격", 5),
                ("4. 자료정의", 6), ("5. 측정기기 점검사항", 7), ("6. 자료생성", 8),
                ("7. 측정기기-자료수집기", 9), ("8. 자료수집기-관제센터", 10)
            ]

            st.markdown("### 📝 1. 통합시험 수행 항목")
            for name, col_idx in test_items:
                if is_checked(target_row.iloc[col_idx]):
                    clean_name = name.replace(" ", "")
                    matched_name = next((val for key, val in sheet_map.items() if key == clean_name), None) or (name if name in report_sheets else None)

                    if matched_name:
                        with st.expander(f"✅ {matched_name}", expanded=True):
                            df_content = report_sheets[matched_name].fillna("")
                            st.dataframe(df_content, use_container_width=True)
                            
                            temp_df = df_content.copy()
                            temp_df.insert(0, '대분류', '통합시험')
                            temp_df.insert(1, '시험항목', matched_name)
                            all_data_frames.append(temp_df)

            # --- 2. 확인검사 항목 ---
            check_items = [
                "외관 및 구조", "전원전압 변동", "절연저항", "공급전압의 안정성", 
                "반복성", "제로 및 스팬 드리프트", "응답시간", "직선성", 
                "유입전류 안정성", "간섭영향", "검출한계"
            ]
            
            st.markdown("### 🔍 2. 확인검사 수행 여부")
            check_list = []
            for i, name in enumerate(check_items):
                status = "수행" if is_checked(target_row.iloc[11 + i]) else "미대상"
                check_list.append({"항목": name, "수행여부": status})
            
            check_df = pd.DataFrame(check_list)
            # 수행해야 할 항목만 화면에 깔끔하게 표시
            active_checks = check_df[check_df["수행여부"] == "수행"]
            if not active_checks.empty:
                st.table(active_checks)
            else:
                st.write("대상 없음")
            
            check_df_for_excel = check_df.copy()
            check_df_for_excel.insert(0, '대분류', '확인검사')
            check_df_for_excel.rename(columns={'항목': '시험항목', '수행여부': '내용/결과'}, inplace=True)
            all_data_frames.append(check_df_for_excel)

            # --- 3. 상대정확도 ---
            st.markdown("### 📊 3. 상대정확도 수행 여부")
            rel_status = "수행 대상" if is_checked(target_row.iloc[22]) else "대상 아님"
            if "수행" in rel_status:
                st.error(f"📍 상대정확도: {rel_status}")
            else:
                st.info(f"📍 상대정확도: {rel_status}")
            
            rel_df = pd.DataFrame([{"대분류": "상대정확도", "시험항목": "상대정확도 시험", "내용/결과": rel_status}])
            all_data_frames.append(rel_df)

            # --- 💾 엑셀 저장 ---
            if all_data_frames:
                st.divider()
                final_combined_df = pd.concat(all_data_frames, ignore_index=True)
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_combined_df.to_excel(writer, index=False, sheet_name='TMS_시험항목_통합')
                    
                    worksheet = writer.sheets['TMS_시험항목_통합']
                    worksheet.set_column(0, 1, 18)
                    worksheet.set_column(2, 10, 25)
                
                st.download_button(
                    label="📥 전체 시험항목 통합 엑셀 다운로드",
                    data=output.getvalue(),
                    file_name=f"TMS_Exam_Items_{selected_sub}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
