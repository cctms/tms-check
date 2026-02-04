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
        # 가이드북 로드
        guide_path = '개선내역에 따른 시험방법(2025 최종).xlsx'
        guide_df = pd.read_excel(guide_path, sheet_name='★최종(가이드북)', skiprows=1)
        guide_df.iloc[:, 1] = guide_df.iloc[:, 1].ffill() # 대분류 채우기
        
        # 통합시험 조사표 로드 (모든 시트 가져오기)
        report_path = '1.통합시험 조사표.xlsx'
        all_sheets = pd.read_excel(report_path, sheet_name=None)
        
        # 시트명 매칭을 위한 전처리 (공백 제거 맵)
        sheet_map = {name.replace(" ", ""): name for name in all_sheets.keys()}
        
        return guide_df, all_sheets, sheet_map
    except Exception as e:
        st.error(f"⚠️ 엑셀 파일을 불러오지 못했습니다. 파일명을 확인하세요: {e}")
        return None, None, None

guide_df, report_sheets, sheet_map = load_all_data()

# 체크 표시 판별 함수
def is_checked(value):
    if pd.isna(value): return False
    val_str = str(value).replace(" ", "").upper()
    return any(m in val_str for m in ['O', '○', '오', 'ㅇ', 'V'])

if guide_df is not None:
    # --- 사이드바: 개선내역 선택 ---
    with st.sidebar:
        st.header("🔍 개선내역 선택")
        categories = guide_df.iloc[:, 1].dropna().unique()
        selected_cat = st.selectbox("1. 대분류", categories)
        
        filtered_df = guide_df[guide_df.iloc[:, 1] == selected_cat]
        sub_items = [str(item).replace('\n', ' ').strip() for item in filtered_df.iloc[:, 2].dropna().unique()]
        selected_sub = st.selectbox("2. 상세내역", ["선택 안 함"] + sub_items)

    if selected_sub != "선택 안 함":
        # 선택된 행 찾기
        target_row = None
        for _, row in filtered_df.iterrows():
            if str(row.iloc[2]).replace('\n', ' ').strip() == selected_sub:
                target_row = row
                break

        if target_row is not None:
            st.success(f"🎯 **선택 내역:** {selected_sub}")
            
            # 매칭할 시험 항목 정의 (가이드북 열 순서 기준)
            test_items = [
                ("1. 일반현황", 3), ("2. 하드웨어 규격", 4), ("3. 소프트웨어 기능 규격", 5),
                ("4. 자료정의", 6), ("5. 측정기기 점검사항", 7), ("6. 자료생성", 8),
                ("7. 측정기기-자료수집기", 9), ("8. 자료수집기-관제센터", 10)
            ]

            final_dfs = [] # 엑셀 저장용 리스트

            st.markdown("### 📝 수행해야 할 통합시험 항목")
            
            # 2단 레이아웃 (왼쪽: 상세 내용 표출, 오른쪽: 요약 정보)
            col_main, col_side = st.columns([2, 1])

            with col_main:
                for name, col_idx in test_items:
                    # 가이드북 해당 열에 체크(O)가 되어 있는지 확인
                    if is_checked(target_row.iloc[col_idx]):
                        clean_name = name.replace(" ", "")
                        matched_name = None
                        
                        # 시트 이름 매칭 시도
                        if name in report_sheets:
                            matched_name = name
                        elif clean_name in sheet_map:
                            matched_name = sheet_map[clean_name]

                        if matched_name:
                            with st.expander(f"✅ {matched_name} 상세 내용", expanded=True):
                                df_content = report_sheets[matched_name]
                                # 데이터 표시 (NaN 제거 및 깔끔하게 출력)
                                display_df = df_content.dropna(how='all').reset_index(drop=True)
                                st.dataframe(display_df, use_container_width=True)
                                
                                # 병합용 리스트에 추가 (시트명 정보 포함)
                                display_df._sheet_name = matched_name
                                final_dfs.append(display_df)
                        else:
                            st.warning(f"⚠️ '{name}'에 해당하는 시트를 엑셀에서 찾을 수 없습니다.")
                    else:
                        st.write(f"⚪ {name}: 대상 아님")

            with col_side:
                # 확인검사 및 상대정확도 요약
                st.markdown("#### 🔍 추가 확인사항")
                # 가이드북 열 11~21번(확인검사) 처리 로직 생략 가능하나 필요시 추가
                if is_checked(target_row.iloc[22]):
                    st.error("📊 상대정확도: **수행 대상**")
                else:
                    st.success("📊 상대정확도: **대상 아님**")

            # --- 통합 엑셀 다운로드 생성 ---
            if final_dfs:
    st.divider()
    output = BytesIO()
    # 엔진을 openpyxl로 변경하여 더 안정적으로 저장합니다.
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for df in final_dfs:
            # 시트명 제한 대응 (31자)
            s_name = str(df._sheet_name)[:31]
            # 파일이 깨지지 않도록 index=False 설정
            df.to_excel(writer, index=False, sheet_name=s_name)
    
    # 중요: 포인터를 처음으로 돌려야 파일 내용이 제대로 전달됩니다.
    data = output.getvalue()
    
    st.download_button(
        label=f"📥 {selected_sub} 통합 조사표 다운로드",
        data=data,
        file_name=f"TMS_Result.xlsx", # 파일명을 일단 간단하게 해서 테스트해보세요
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
                
                st.download_button(
                    label=f"📥 {selected_sub} 통합 조사표 다운로드",
                    data=output.getvalue(),
                    file_name=f"TMS_통합시험_{selected_sub}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.info("왼쪽 사이드바에서 개선내역을 선택하면 해당되는 통합시험 조사표를 발췌합니다.")

