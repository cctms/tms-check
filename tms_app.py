import streamlit as st
import pandas as pd

# 페이지 설정
st.set_page_config(page_title="TMS 통합조사표 생성기", layout="wide")

st.title("📊 TMS 통합조사표 생성기")
st.write("엑셀 파일을 업로드하여 내용을 확인하고 관리하세요.")

# 파일 업로드 창
uploaded_file = st.file_uploader("📂 엑셀 파일을 업로드하세요 (xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # 엑셀 읽기
    df = pd.read_excel(uploaded_file)
    
    # 상단에 간단한 통계 보여주기
    col1, col2, col3 = st.columns(3)
    col1.metric("전체 항목 수", f"{len(df)}개")
    
    # 데이터 표 출력
    st.subheader("📋 조사표 데이터 미리보기")
    st.dataframe(df, use_container_width=True)
    
    st.success("✅ 파일을 성공적으로 불러왔습니다!")
else:
    st.info("💡 엑셀 파일을 업로드하면 이곳에 표가 나타납니다.")
