import streamlit as st
import pandas as pd
from io import BytesIO
import os

st.set_page_config(page_title="TMS 수질 시험 매칭", layout="wide")

@st.cache_data
def load_all_data():
    try:
        f_list = os.listdir('.')
        g_p = next((f for f in f_list if '가이드북' in f or '시험방법' in f), None)
        r_p = next((f for f in f_list if '1.통합' in f), None)
        c_p = next((f for f in f_list if '2.확인' in f), None)
        s_p = next((f for f in f_list if '상대' in f or '3.' in f), None)
        if not g_p: return None, None, None, None
        
        # 가이드북 로드 및 실제 헤더 위치(항목명 행) 찾기
        guide_raw = pd.read_excel(g_p, header=None)
        header_idx = 2
        for i in range(min(6, len(guide_raw))):
            row_str = "".join(guide_raw.iloc[i].astype(str))
            if any(k in row_str for k in ["반복성", "제로드리프트", "일반현황"]):
                header_idx = i
                break
        
        df_g = pd.read_excel(g_p, skiprows=header_idx)
        df_g.iloc[:, 1] = df_g.iloc[:, 1].ffill()
        
        r_s = pd.read_excel(r_p, sheet_name=None) if r_p else {}
        c_s = pd.read_excel(c_p, sheet_name=None) if c_p else {}
        s_s = pd.read_excel(s_p, sheet_name=None) if s_p else {}
        
        return df_g, r_s, c_s, s_s
    except Exception as e:
        return None, None, None, None

df_g, r_s, c_s, s_s = load_all_data()

def is_ok(v):
    if pd.isna(v): return False
    s = str(v).replace(" ", "").upper()
    return any(m in s for m in ['O', 'ㅇ', '○', 'V', '◎', '대상'])

st.title("📋 수질 TMS 개선내역 매칭 결과")

if df_g is not None:
    search_q = st.text_input("개선내역 입력 (예: 측정기기 교체)", "")
    if search_q:
        match_rows = df_g[df_g.iloc[:, 2].astype(str).str.contains(search_q, na=False)]
        if not match_rows.empty:
            match_rows['dn'] = match_rows.apply(lambda x: f"[{x.iloc[1]}] {x.iloc[2]}", axis=1)
            sel = st.selectbox("항목 선택", ["선택하세요"] + match_rows['dn'].tolist())
            
            if sel != "선택하세요":
                row = match_rows[match_rows['dn'] == sel].iloc[0]
                
                # 엑셀 열 순서(index)를 유지한 채로 체크된 시험명 리스트 생성
                active_list = []
                for col in df_g.columns:
                    if is_ok(row[col]):
                        c_name = str(col).strip()
                        if not any(ex in c_name for ex in ["순번", "분류", "개선내역", "Unnamed"]):
                            active_list.append(c_name)

                # 파일별 분류 키워드 (통합/확인/상대 구분을 위해)
                r_keys = ["일반현황", "점검사항", "자료생성", "자료수집기", "관제센터"]
                c_keys = ["구조", "시료", "승인", "방법", "범위", "교정", "표준물질", "정도검사", "교정일자", "유량계", "누적값", "반복성", "드리프트", "재현성"]

                def find_tabs(test_name, sheets, f_type):
                    matched = []
                    tn = test_name.replace(" ", "")
                    for sn in sheets.keys():
                        sn_c = str(sn).replace(" ", "")
                        # 직접 포함 매칭
                        if tn in sn_c or sn_c in tn: matched.append(sn)
                        # 특수 규칙 (외관/구조/점검 등)
                        elif f_type == "확인" and any(k in tn for k in ["외관", "구조"]):
                            if any(k in sn_c for k in ["구조", "시료", "승인", "방법", "범위", "교정", "일자"]): matched.append(sn)
                        elif f_type == "통합" and any(k in tn for k in ["점검", "생성", "전송"]):
                            if any(k in sn_c for k in ["점검", "생성", "전송", "관제"]): matched.append(sn)
                    return list(dict.fromkeys(matched)) # 중복 제거 유지

                all_data = []
                col1, col2, col3 = st.columns(3)

                with col1:
                    st.header("1. 통합시험")
                    for t_name in active_list:
                        if any(k in t_name for k in r_keys):
                            st.subheader(f"📍 {t_name}")
                            tabs = find_tabs(t_name, r_s, "통합")
                            for t in tabs:
                                with st.expander(f"📑 {t}"):
                                    st.dataframe(r_s[t].fillna(""))
                                    tmp = r_s[t].copy(); tmp.insert(0, '시험항목', t_name); tmp.insert(1, '탭이름', t); all_data.append(tmp)

                with col2:
                    st.header("2. 확인검사")
                    for t_name in active_list:
                        if any(k in t_name for k in c_keys):
                            st.subheader(f"📍 {t_name}")
                            tabs = find_tabs(t_name, c_s, "확인")
                            for t in tabs:
                                with st.expander(f"📑 {t}"):
                                    st.dataframe(c_s[t].fillna(""))
                                    tmp = c_s[t].copy(); tmp.insert(0, '시험항목', t_name); tmp.insert(1, '탭이름', t); all_data.append(tmp)

                with col3:
                    st.header("3. 상대정확도")
                    for t_name in active_list:
                        if "상대" in t_name:
                            st.subheader(f"📍 {t_name}")
                            for t in s_s.keys():
                                with st.expander(f"📑 {t}"):
                                    st.dataframe(s_s[t].fillna(""))
                                    tmp = s_s[t].copy(); tmp.insert(0, '시험항목', t_name); tmp.insert(1, '탭이름', t); all_data.append(tmp)

                if all_data:
                    out = BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
                        pd.concat(all_data).to_excel(wr, index=False)
                    st.download_button("📥 통합 리포트 다운로드", out.getvalue(), "TMS_Matching_Full.xlsx")
