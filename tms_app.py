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
        guide_df.iloc[:, 1] = guide_df.iloc[:, 1].ffill()
        
        # 통합시험 조사표 로드
        report_path = '1.통합시험 조사표.xlsx'
        all_sheets = pd.read_excel(report_path, sheet_name=None)
        
        # 시트명 매칭용 맵 (공백 제거)
        sheet_map = {name.replace(" ", ""): name for name in all_sheets.keys()}
        
        return guide_df, all_sheets, sheet_map
    except Exception as e:
        st.error(f"⚠️ 엑셀 파일을 불러오지 못했습니다: {e}")
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
        target_row = None
        for _, row in filtered_df.iterrows():
            if str(row.iloc[2]).replace('\n', ' ').strip() == selected_sub:
                target_row = row
                break

        if target_row is not None:
            st.success(f"🎯 **선택 내역:** {selected_sub}")
            
            # 매칭할 시험 항목 (가이드북 열 순서 기준)
            test_items = [
                ("1. 일반현황", 3), ("2. 하드웨어 규격", 4), ("3. 소프트웨어 기능 규격", 5),
                ("4. 자료정의", 6), ("5. 측정기기 점검사항", 7), ("6. 자료생성", 8),
                ("7. 측정기기-자료수집기", 9), ("8. 자료수집기-관제센터", 10)
            ]

            final_dfs = [] # 엑셀 병합용 리스트

            st.markdown("### 📝 수행해야 할 통합시험 항목")
            col_main, col_side = st.columns([2, 1])

            with col_main:
                for name, col_idx in test_items:
                    if is_checked(target_row.iloc[col_idx]):
                        clean_name = name.replace(" ", "")
                        matched_name = None
                        
                        if name in report_sheets:
                            matched_name = name
                        elif clean_name in sheet_map:
                            matched_name = sheet_map[clean_name]

                        if matched_name:
                            with st.expander(f"✅ {matched_name} 상세 내용", expanded=True):
                                df_content = report_sheets[matched_name]
                                display_df = df_content.dropna(how='all').reset_index(drop=True)
                                st.dataframe(display_df, use_container_width=True)
                                
                                # 엑셀 저장을 위해 시트명을 데이터프레임 속성으로 임시 저장
                                # 복사본을 만들어 데이터 오염 방지
                                excel_df = display_df.copy()
                                excel_df.attrs['sheet_name'] = matched_name
                                final_dfs.append(excel_df)
                        else:
                            st.warning(f"⚠️ '{name}' 시트를 엑셀에서 찾을 수 없습니다.")
                    else:
                        st.write(f"⚪ {name}: 대상 아님")

            with col_side:
                st.markdown("#### 🔍 추가 확인사항")
                if is_checked(target_row.iloc[22]):
                    st.error("📊 상대정확도: **수행 대상**")
                else:
                    st.success("📊 상대정확도: **대상 아님**")

            # --- 통합 엑셀 다운로드 생성 (오류 방지 로직 강화) ---
            if final_dfs:
                st.divider()
                
                # 메모리 버퍼에 엑셀 파일 생성
                output = BytesIO()
                try:
                    # engine='openpyxl'이 가장 호환성이 좋습니다.
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        for df in final_dfs:
                            # 엑셀 시트 이름 규칙: 최대 31자, 특수문자 / \ ? * : [ ] 제한
                            s_name = df.attrs.get('sheet_name', 'Sheet')
                            s_name = "".join([c for c in s_name if c not in r'/\?*:[]'])[:31]
                            df.to_excel(writer, index=False, sheet_name=s_name)
                    
                    # 데이터 준비
                    excel_data = output.getvalue()
                    
                    st.download_button(
                        label=f"📥 {selected_sub} 통합 조사표 다운로드",
                        data=excel_data,
                        file_name=f"TMS_Checklist_{selected_sub}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"엑셀 파일 생성 중 오류가 발생했습니다: {e}")
    else:
        st.info("왼쪽 사이드바에서 상세내역을 선택해 주세요.")
