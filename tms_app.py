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
        # 파일명 확인 필수! (GitHub에 올린 이름과 대소문자까지 같아야 합니다)
        guide_path = '개선내역에 따른 시험방법(2025 최종).xlsx'
        report_path = '1.통합시험 조사표.xlsx'
        
        guide_df = pd.read_excel(guide_path, sheet_name='★최종(가이드북)', skiprows=1)
        guide_df.iloc[:, 1] = guide_df.iloc[:, 1].ffill()
        
        report_sheets = pd.read_excel(report_path, sheet_name=None)
        sheet_map = {name.replace(" ", ""): name for name in report_sheets.keys()}
        
        return guide_df, report_sheets, sheet_map
    except Exception as e:
        st.error(f"⚠️ 파일을 찾을 수 없습니다. GitHub 저장소에 엑셀 파일이 있는지 확인하세요. ({e})")
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
                                df_content = report_sheets[matched_name].dropna(how='all').reset_index(drop=True)
                                st.dataframe(df_content, use_container_width=True)
                                # 시트 이름을 데이터에 직접 박아넣지 않고 별도 저장
                                final_dfs.append((matched_name, df_content))
                        else:
                            st.warning(f"⚠️ '{name}' 시트 없음")

            with col_side:
                st.markdown("#### 🔍 추가 확인")
                if is_checked(target_row.iloc[22]): st.error("📊 상대정확도: **대상**")
                else: st.success("📊 상대정확도: **미대상**")

            # --- 🔥 오류 해결 핵심: 다운로드 로직 ---
            if final_dfs:
                st.divider()
                
                # 1. 메모리에 엑셀 파일 생성
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for s_name, df in final_dfs:
                        # 엑셀 시트명 규칙 적용 (특수문자 제거 및 31자 제한)
                        safe_name = "".join([c for c in s_name if c not in r'/\?*:[]'])[:31]
                        df.to_excel(writer, index=False, sheet_name=safe_name)
                
                # 2. 버퍼의 포인터를 0으로 돌려야 데이터가 누락되지 않음
                output.seek(0)
                processed_data = output.getvalue()
                
                # 3. 데이터가 비어있는지 확인 후 버튼 생성
                if processed_data:
                    st.download_button(
                        label="📥 발췌된 조사표 다운로드 (Excel)",
                        data=processed_data,
                        file_name=f"TMS_Result.xlsx", # 한글명 에러 방지를 위해 영어로 설정
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
