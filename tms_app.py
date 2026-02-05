카테고리를 일일이 클릭하는 방식 대신, 질문창(검색창)에 키워드를 입력하면 관련 개선내역을 찾아서 시험항목을 보여주는 방식으로 업그레이드해 보겠습니다.

사용자가 입력한 단어가 포함된 모든 개선내역을 찾아주고, 그중 하나를 선택하면 기존처럼 상세 시험항목을 출력하는 구조입니다.

🚀 검색 기능이 추가된 버전 (app.py)
Python
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
        # 전개된 카테고리(대분류) 채우기
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
    # --- 🔍 검색 기능 추가 ---
    st.markdown("### 🔍 개선내역 검색")
    search_query = st.text_input("찾으시는 개선내역의 키워드를 입력하세요 (예: 전송, 통신, 부착)", "")

    # 검색어에 맞는 데이터 필터링
    if search_query:
        # 2번 열(상세내역)에서 검색어 포함 여부 확인
        search_results = guide_df[guide_df.iloc[:, 2].str.contains(search_query, na=False, case=False)]
        
        if not search_results.empty:
            # 검색된 항목들을 리스트로 보여주고 선택하게 함
            options = [f"[{row.iloc[1]}] {str(row.iloc[2]).strip()}" for _, row in search_results.iterrows()]
            selected_option = st.selectbox(f"검색 결과 ({len(options)}건):", ["선택하세요"] + options)
            
            if selected_option != "선택하세요":
                # 선택된 항목의 실제 데이터 행 추출
                idx = options.index(selected_option)
                target_row = search_results.iloc[idx]
                selected_sub = str(target_row.iloc[2]).replace('\n', ' ').strip()
                
                # --- 상세 항목 출력 로직 시작 ---
                st.divider()
                st.success(f"🎯 **분석 결과:** {selected_sub}")
                
                all_data_frames = []

                # 1. 통합시험 항목
                test_items = [
                    ("1. 일반현황", 3), ("2. 하드웨어 규격", 4), ("3. 소프트웨어 기능 규격", 5),
                    ("4. 자료정의", 6), ("5. 측정기기 점검사항", 7), ("6. 자료생성", 8),
                    ("7. 측정기기-자료수집기", 9), ("8. 자료수집기-관제센터", 10)
                ]

                st.markdown("### 📝 1. 통합시험 수행 항목")
                cols = st.columns(2)
                for i, (name, col_idx) in enumerate(test_items):
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

                # 2. 확인검사 및 상대정확도 처리
                st.divider()
                c1, c2 = st.columns(2)
                
                with c1:
                    st.markdown("### 🔍 2. 확인검사")
                    check_items = ["외관 및 구조", "전원전압 변동", "절연저항", "공급전압의 안정성", "반복성", "제로 및 스팬 드리프트", "응답시간", "직선성", "유입전류 안정성", "간섭영향", "검출한계"]
                    check_list = []
                    for i, name in enumerate(check_items):
                        status = "수행" if is_checked(target_row.iloc[11 + i]) else "미대상"
                        check_list.append({"항목": name, "수행여부": status})
                    
                    active_checks = pd.DataFrame(check_list)
                    st.table(active_checks[active_checks["수행여부"] == "수행"])
                    
                    check_df_for_excel = active_checks.copy()
                    check_df_for_excel.insert(0, '대분류', '확인검사')
                    check_df_for_excel.rename(columns={'항목': '시험항목', '수행여부': '내용/결과'}, inplace=True)
                    all_data_frames.append(check_df_for_excel)

                with c2:
                    st.markdown("### 📊 3. 상대정확도")
                    rel_status = "수행 대상" if is_checked(target_row.iloc[22]) else "대상 아님"
                    if "수행" in rel_status:
                        st.error(f"📍 {rel_status}")
                    else:
                        st.info(f"📍 {rel_status}")
                    
                    rel_df = pd.DataFrame([{"대분류": "상대정확도", "시험항목": "상대정확도 시험", "내용/결과": rel_status}])
                    all_data_frames.append(rel_df)

                # 3. 엑셀 다운로드
                if all_data_frames:
                    st.divider()
                    final_df = pd.concat(all_data_frames, ignore_index=True)
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        final_df.to_excel(writer, index=False, sheet_name='전체항목')
                    
                    st.download_button(
                        label="📥 결과 엑셀 다운로드",
                        data=output.getvalue(),
                        file_name=f"TMS_Search_{search_query}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            st.warning("검색 결과가 없습니다. 다른 키워드를 입력해 보세요.")
    else:
        st.info("검색창에 개선내역 키워드를 입력하시면 관련 항목을 찾아드립니다.")
