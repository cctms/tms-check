import streamlit as st
import pandas as pd
from io import BytesIO
import os

# 1. 페이지 설정
st.set_page_config(page_title="수질 TMS 시험항목 도구", layout="wide")

# 스타일 설정: 분석 결과 줄바꿈 방지
st.markdown("""
    <style>
    .single-line-header {
        white-space: nowrap;
        overflow-x: auto;
        font-size: 1.6rem;
        font-weight: 700;
        padding: 10px 0px;
        color: #0E1117;
        border-bottom: 2px solid #F0F2F6;
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("📋 수질 TMS 개선내역별 시험항목")

# 2. 데이터 로드 함수
@st.cache_data
def load_all_data():
    try:
        files = os.listdir('.')
        guide_path = next((f for f in files if '가이드북' in f or '시험방법' in f), None)
        report_path = next((f for f in files if '1.통합시험' in f), None)
        check_path = next((f for f in files if '2.확인검사' in f), None)
        rel_path = next((f for f in files if '상대정확도' in f or '3.상대정확도' in f), None)
        
        if not guide_path:
            st.error(f"❌ 가이드북 파일을 찾을 수 없습니다. 현재 폴더 파일: {files}")
            return None, None, None, None
            
        guide_df = pd.read_excel(guide_path, sheet_name='★최종(가이드북)', skiprows=1)
        guide_df.iloc[:, 1] = guide_df.iloc[:, 1].ffill()
        
        report_sheets = pd.read_excel(report_path, sheet_name=None) if report_path else {}
        check_sheets = pd.read_excel(check_path, sheet_name=None) if check_path else {}
        rel_sheets = pd.read_excel(rel_path, sheet_name=None) if rel_path else {}
        
        return guide_df, report_sheets, check_sheets, rel_sheets
    except Exception as e:
        st.error(f"⚠️ 데이터 로드 중 오류 발생: {e}")
        return None, None
