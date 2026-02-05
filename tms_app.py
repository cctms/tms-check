import streamlit as st
import pandas as pd
from io import BytesIO

# 1. 페이지 설정
st.set_page_config(page_title="TMS 시험항목 도구", layout="wide")

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
    st.markdown("### 🔍 개선내역 검색")
    search_query = st.text_input("찾으시는 개선내역의 키워드를 입력하세요", "")

    if search_query:
        search_results = guide_df[guide_df.iloc[:, 2].str.contains(search_query, na=False, case=False)].copy()
        
        if not search_results.empty:
            search_results['display_name'] = search_results.apply(lambda x: f"[{x.iloc[1]}] {str(x.iloc[2]).strip()}", axis=1)
            options = search_results['display_name'].tolist()
            
            selected_option = st.selectbox(f"검색 결과 ({len(options)}건):", ["선택하세요"] + options)
            
            if selected_option != "선택하세요":
                target_row = search_results[search_results['display_name'] == selected_option].iloc[0]
                full_display_name = selected_option 
                selected_sub = str(target_row.iloc[2]).replace('\n', ' ').strip()
                
                st.divider()
                
                # 제목 줄바꿈 방지 스타일
                st.markdown(
                    f"""
                    <div style="white-space: nowrap; overflow-x: auto; font-size: 1.6rem; font-weight: 700; 
                    padding: 10px 0px; color: #0E1117; border-bottom: 2px solid #F0F2F6; margin-bottom: 20px;">
                        🎯 분석 결과: {full_display_name}
                    </div>
                    """, 
                    unsafe_allow_html=True
                )
                
                all_data_frames = []

                # --- 🎨 3단 레이아웃 (너비 동일하게 [1, 1, 1]) ---
                col1, col2, col3 = st.columns([1, 1, 1])

                # [1단: 통합시험]
                with col1:
                    st.markdown("#### 📝 1. 통합시험")
                    test_items = [
                        ("1. 일반현황", 3), ("2. 하드웨어 규격", 4), ("3. 소프트웨어 기능 규격", 5),
                        ("4. 자료정의", 6), ("5. 측정기기 점검사항", 7), ("6. 자료생성", 8),
                        ("7. 측정기기-자료수집기", 9), ("8. 자료수집기-관제센터", 10)
                    ]
                    found_test = False
                    for name, col_idx in test_items:
                        if is_checked(target_row.iloc[col_idx]):
                            found_test = True
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
                    
                    if not found_test:
                        st.info("📍 대상 아님")

                # [2단: 확인검사]
                with col2:
                    st.markdown("#### 🔍 2. 확인검사")
                    check_items = ["외관 및 구조", "전원전압 변동", "절연저항", "공급전압의 안정성", "반복성", "제로 및 스팬 드리프트", "응답시간", "직선성", "유입전류 안정성", "간섭영향", "검출한계"]
                    check_list = []
                    for i, name in enumerate(check_items):
                        if is_checked(target_row.iloc[11 + i]):
                            check_list.append({"항목": name, "수행여부": "수행"})
                    
                    if check_list:
                        active_checks = pd.DataFrame(check_list)
                        st.table(active_checks)
                        check_df_excel = active_checks.copy()
                        check_df_excel.insert(0, '대분류', '확인검사')
                        check_df_excel.rename(columns={'항목': '시험항목', '수행여부': '내용/결과'}, inplace=True)
                        all_data_frames.append(check_df_excel)
                    else:
                        st.info("📍 대상 아님")

                # [3단: 상대정확도]
                with col3:
                    st.markdown("#### 📊 3. 상대정확도")
                    rel_status = "수행 대상" if is_checked(target_row.iloc[22]) else "대상 아님"
                    if "수행" in rel_status:
                        st.error(f"📍 {rel_status}")
                    else:
                        st.info(f"📍 {rel_status}")
                    
                    rel_df = pd.DataFrame([{"대분류": "상대정확도", "시험항목": "상대정확도 시험", "내용/결과": rel_status}])
                    all_data_frames.append(rel_df)

                # --- 💾 엑셀 다운로드 ---
                if all_data_frames:
                    st.divider()
                    final_df = pd.concat(all_data_frames, ignore_index=True)
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        final_df.to_excel(writer, index=False, sheet_name='전체항목')
                    
                    st.download_button(
                        label="📥 전체 결과 엑셀 다운로드",
                        data=output.getvalue(),
                        file_name=f"TMS_Report_{selected_sub}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            st.warning("검색 결과가 없습니다.")
    else:
        st.info("검색창에 개선내역 키워드를 입력하세요.")
