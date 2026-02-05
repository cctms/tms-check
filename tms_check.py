import streamlit as st
import pandas as pd
from io import BytesIO
import os

# 1. 페이지 설정
st.set_page_config(page_title="수질 TMS 시험항목 도구", layout="wide")

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
        if not guide_path: return None, None, None, None
        
        guide_df = pd.read_excel(guide_path, sheet_name='★최종(가이드북)', skiprows=1)
        guide_df.iloc[:, 1] = guide_df.iloc[:, 1].ffill()
        
        report_sheets = pd.read_excel(report_path, sheet_name=None) if report_path else {}
        check_sheets = pd.read_excel(check_path, sheet_name=None) if check_path else {}
        rel_sheets = pd.read_excel(rel_path, sheet_name=None) if rel_path else {}
        
        return guide_df, report_sheets, check_sheets, rel_sheets
    except:
        return None, None, None, None

guide_df, report_sheets, check_sheets, rel_sheets = load_all_data()

def is_checked(value):
    if pd.isna(value): return False
    v = str(value).replace(" ", "").upper()
    return any(m in v for m in ['O', '○', '오', 'ㅇ', 'V', 'CHECK'])

if guide_df is not None:
    st.markdown("### 🔍 개선내역 검색")
    query = st.text_input("키워드를 입력하세요 (예: 기기교체)", "")

    if query:
        res = guide_df[guide_df.iloc[:, 2].str.contains(query, na=False, case=False)].copy()
        if not res.empty:
            res['d_name'] = res.apply(lambda x: f"[{x.iloc[1]}] {str(x.iloc[2]).strip()}", axis=1)
            sel = st.selectbox(f"검색 결과 ({len(res)}건):", ["선택하세요"] + res['d_name'].tolist())
            
            if sel != "선택하세요":
                row = res[res['d_name'] == sel].iloc[0]
                sub_text = str(row.iloc[2]).replace('\n', ' ').strip()
                all_dfs = []
                col1, col2, col3 = st.columns([1, 1, 1])

                # [1. 통합시험]
                with col1:
                    st.markdown("#### 📝 1. 통합시험")
                    t_items = [
                        ("1. 일반현황", 3), ("2. 하드웨어 규격", 4), 
                        ("3. 소프트웨어 기능 규격", 5), ("4. 자료정의", 6), 
                        ("5. 측정기기 점검사항", 7), ("6. 자료생성", 8), 
                        ("7. 측정기기-자료수집기", 9), ("8. 자료수집기-관제센터", 10)
                    ]
                    is_교체 = "교체" in sub_text
                    f_test = any(is_checked(row.iloc[idx]) for _, idx in t_items) or is_교체
                    
                    if f_test:
                        st.error("📍 수행 대상")
                        for name, idx in t_items:
                            if is_checked(row.iloc[idx]) or (is_교체 and idx in [9, 10]):
                                # 시트 찾기 로직
                                m_name = next((s for s in report_sheets.keys() if s.strip() == name.strip()), None)
                                if not m_name:
                                    pref = name.split('.')[0] + "."
                                    m_name = next((s for s in report_sheets.keys() if s.strip().startswith(pref)), None)
                                
                                if m_name:
                                    with st.expander(f"✅ {name}"):
                                        df = report_sheets[m_name].fillna("")
                                        st.dataframe(df, use_container_width=True)
                                        df_exp = df.copy()
                                        df_exp.insert(0, '대분류', '통합시험')
                                        df_exp.insert(1, '시험항목', name)
                                        all_dfs.append(df_exp)
                                else:
                                    st.warning(f"⚠️ {name} (연결 실패)")

                # [2. 확인검사]
                with col2:
                    st.markdown("#### 🔍 2. 확인검사")
                    c_names = ["외관 및 구조", "전원전압 변동", "절연저항", "공급전압의 안정성", "반복성", "제로 및 스팬 드리프트", "응답시간", "직선성", "유입전류 안정성", "간섭영향", "검출한계"]
                    w_sheets = ["측정소 구조 및 설비", "시료채취조", "형식승인", "측정방법", "측정범위", "교정기능(표준물질)", "정도검사 교정일자"]
                    
                    f_check = any(is_checked(row.iloc[11 + i]) for i in range(len(c_names)))
                    if f_check:
                        st.error("📍 수행 대상")
                        for i, name in enumerate(c_names):
                            if is_checked(row.iloc[11 + i]):
                                if name == "외관 및 구조":
                                    for sn in w_sheets:
                                        if sn in check_sheets:
                                            with st.expander(f"✅ {sn}"):
                                                df = check_sheets[sn].fillna("")
                                                st.dataframe(df, use_container_width=True)
                                                df_e = df.copy(); df_e.insert(0, '대분류', '확인검사'); df_e.insert(1, '시험항목', sn); all_dfs.append(df_e)
                                elif name in check_sheets:
                                    with st.expander(f"✅ {name}"):
                                        df = check_sheets[name].fillna("")
                                        st.dataframe(df, use_container_width=True)
                                        df_e = df.copy(); df_e.insert(0, '대분류', '확인검사'); df_e.insert(1, '시험항목', name); all_dfs.append(df_e)
                                else: st.write(f"✅ {name}")

                # [3. 상대정확도]
                with col3:
                    st.markdown("#### 📊 3. 상대정확도")
                    if is_checked(row.iloc[22]):
                        st.error("📍 수행 대상")
                        if rel_sheets:
                            r_n = next((s for s in rel_sheets.keys() if '상대정확도' in s), list(rel_sheets.keys())[0])
                            with st.expander("✅ 상대정확도 결과서"):
                                df = rel_sheets
