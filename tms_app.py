import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="TMS 통합조사표 생성기", layout="wide")

st.title("🚀 TMS 통합조사표 자동 생성기")
st.write("개선내역 파일과 조사표 파일을 업로드하면 데이터를 하나로 합쳐줍니다.")

# 1. 파일 업로드 섹션
col1, col2 = st.columns(2)
with col1:
    file_method = st.file_uploader("📂 개선내역별 시험방법 업로드", type=["xlsx"])
with col2:
    file_survey = st.file_uploader("📂 통합시험 조사표 양식 업로드", type=["xlsx"])

if file_method and file_survey:
    try:
        # 데이터 로드
        df_method = pd.read_excel(file_method)
        df_survey = pd.read_excel(file_survey)

        st.subheader("📊 데이터 병합 처리")
        
        # 2. 데이터 병합 (두 파일의 공통 컬럼인 '시험항목' 기준)
        # ※ 실제 엑셀의 컬럼명에 따라 '시험항목' 부분을 수정해야 할 수 있습니다.
        merged_df = pd.merge(df_method, df_survey, on="시험항목", how="inner")

        if not merged_df.empty:
            st.success(f"✅ 총 {len(merged_df)}건의 매칭된 데이터를 찾았습니다!")
            st.dataframe(merged_df, use_container_width=True)

            # 3. 엑셀 다운로드 버튼
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                merged_df.to_excel(writer, index=False, sheet_name='통합조사표_결과')
            
            st.download_button(
                label="📥 합쳐진 엑셀 파일 다운로드",
                data=output.getvalue(),
                file_name="TMS_통합조사표_결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ 두 파일에서 일치하는 '시험항목'을 찾지 못했습니다. 컬럼명을 확인해 주세요.")

    except Exception as e:
        st.error(f"❌ 오류 발생: {e}")
else:
    st.info("💡 왼쪽에는 개선내역 파일을, 오른쪽에는 조사표 양식 파일을 업로드해 주세요.")
