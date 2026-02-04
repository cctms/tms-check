import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="TMS 통합조사표 생성기", layout="wide")

st.title("🚀 TMS 개선내역별 상세항목 발췌 도구")
st.info("개선내역을 선택하면 해당되는 시험방법 항목만 발췌하여 엑셀로 저장합니다.")

# 1. 파일 업로드
col1, col2 = st.columns(2)
with col1:
    file_method = st.file_uploader("📂 개선내역별 시험방법 엑셀 업로드", type=["xlsx"])
with col2:
    file_survey = st.file_uploader("📂 통합시험 조사표 엑셀 업로드", type=["xlsx"])

if file_method and file_survey:
    # 데이터 로드
    df_method = pd.read_excel(file_method)
    df_survey = pd.read_excel(file_survey)

    # 2. 개선내역 선택 (사용자가 고를 수 있게)
    target_list = df_method['개선내역'].unique()
    selected_target = st.selectbox("🎯 발췌할 개선내역을 선택하세요", target_list)

    if selected_target:
        # 3. 데이터 필터링 (해당 개선내역의 시험방법 찾기)
        # 예: 선택한 개선내역에 해당하는 '시험항목'이나 'ID'를 기준으로 발췌
        target_methods = df_method[df_method['개선내역'] == selected_target]
        
        # 통합시험 조사표에서 해당 항목들만 추출
        # (조사표의 '시험항목' 컬럼이 기준이라고 가정)
        result_df = df_survey[df_survey['시험항목'].isin(target_methods['시험항목'])]

        st.success(f"✅ '{selected_target}'에 해당하는 {len(result_df)}개의 항목을 찾았습니다.")
        st.dataframe(result_df) # 화면에 미리보기

        # 4. 엑셀 파일 생성 및 다운로드
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='발췌내역')
        
        processed_data = output.getvalue()

        st.download_button(
            label="📥 발췌된 엑셀 파일 다운로드",
            data=processed_data,
            file_name=f"{selected_target}_상세내역.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.warning("먼저 두 개의 엑셀 파일을 모두 업로드해 주세요.")
