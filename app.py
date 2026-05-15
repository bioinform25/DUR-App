import streamlit as st
import pandas as pd
from itertools import combinations
import re

# --- 1. 페이지 설정 ---
st.set_page_config(page_title="종합 DUR 분석 시스템", layout="wide")
st.title("🛡️ 스마트 다제약물 DUR 통합 분석기")
st.markdown("심평원 8개 카테고리 DUR 데이터를 기반으로 환자의 복용 안전성을 종합 점검합니다.")

# --- 2. 설정 지도 (파일별 컬럼명 및 스킵할 줄 수 정의) ---
DUR_CONFIG = {
    "병용금기": {"file": "1_interaction.xlsx", "skip": 1, "ing1": "유효성분 '1'", "ing2": "유효성분 '2'", "msg": "상세정보"},
    "연령금기": {"file": "2_age.xlsx", "skip": 1, "ing1": "성분명", "msg": "상세정보", "extra": "연령기준"},
    "임부금기": {"file": "3_pregnancy.xlsx", "skip": 1, "ing1": "성분명", "msg": "상세정보", "extra": "임부금기(등급)"},
    "효능군중복": {"file": "4_duplicate.xlsx", "skip": 2, "ing1": "성분명(영문)", "msg": "비고", "extra": "효능군"},
    "용량주의": {"file": "5_dosage.xlsx", "skip": 1, "ing1": "성분명(영문)", "msg": "비고", "extra": "1일 최대용량"},
    "투여기간": {"file": "6_duration.xlsx", "skip": 1, "ing1": "성분명(영문)", "msg": "비고", "extra": "최대 투여기간"},
    "노인주의": {"file": "7_elderly.xlsx", "skip": 1, "ing1": "성분명(영문)", "msg": "비고"},
    "수유부주의": {"file": "8_lactation.xlsx", "skip": 1, "ing1": "성분명(영문)", "msg": "비고"}
}

# --- 3. 데이터 로드 (캐싱하여 속도 최적화) ---
@st.cache_data
def load_all_data():
    # 1. 약제급여목록 (상품명 -> 성분명 매핑용)
    try:
        mapping_db = pd.read_excel('약제급여목록.xlsx')[['제품명', '주성분명']].dropna()
    except Exception as e:
        st.error(f"'약제급여목록.xlsx' 파일을 읽을 수 없습니다: {e}")
        st.stop()

    # 2. 8개 DUR 데이터 로드
    dur_db = {}
    for category, config in DUR_CONFIG.items():
        try:
            dur_db[category] = pd.read_excel(config["file"], skiprows=config["skip"])
        except Exception as e:
            dur_db[category] = pd.DataFrame() # 파일이 없으면 빈 데이터로 넘김
            
    return mapping_db, dur_db

with st.spinner("데이터베이스를 연동하는 중입니다..."):
    mapping_db, dur_db = load_all_data()

# --- 4. 분석 유틸리티 함수 ---
def get_ingredient(product_name):
    """상품명으로 주성분(영문) 추출"""
    match = mapping_db[mapping_db['제품명'] == product_name]
    return match.iloc[0]['주성분명'] if not match.empty else None

def get_searchable_word(ing_string):
    """성분명에서 매칭에 사용할 핵심 영문 단어 하나만 추출 (예: 'doxepin hcl' -> 'doxepin')"""
    if not ing_string: return ""
    clean = re.sub(r'[^a-zA-Z\s]', ' ', str(ing_string)).strip().lower()
    return clean.split()[0] if clean else ""

def search_in_df(df, col_name, search_word):
    """특정 데이터프레임의 컬럼에서 성분명 검색"""
    if df.empty or col_name not in df.columns or not search_word:
        return pd.DataFrame()
    return df[df[col_name].astype(str).str.lower().str.contains(search_word, na=False)]

# --- 5. 사용자 UI 구성 ---
if "basket" not in st.session_state:
    st.session_state.basket = []

# 왼쪽 사이드바 (약물 검색 및 추가)
with st.sidebar:
    st.header("🛒 약물 추가")
    
    # 검색창 활용 (입력 시 자동완성)
    search_term = st.text_input("상품명 검색 (예: 타이레놀, 포크랄)")
    
    if search_term:
        # 검색어가 포함된 약물 목록 필터링
        search_results = mapping_db[mapping_db['제품명'].str.contains(search_term, na=False, case=False)]
        
        if not search_results.empty:
            selected = st.selectbox("검색 결과 (선택하세요):", options=search_results['제품명'].tolist())
            if st.button("내 약통에 추가", use_container_width=True):
                if selected not in st.session_state.basket:
                    st.session_state.basket.append(selected)
                    st.success(f"{selected} 추가됨!")
                else:
                    st.warning("이미 추가된 약물입니다.")
        else:
            st.error("검색 결과가 없습니다.")
            
    st.divider()
    if st.button("🗑️ 약통 전체 비우기", use_container_width=True):
        st.session_state.basket = []
        st.rerun()

# 메인 화면 (분석 결과)
st.subheader("📋 현재 복용 약물 목록")
if not st.session_state.basket:
    st.info("왼쪽 사이드바에서 약물을 검색하여 추가해 주세요.")
else:
    # 담긴 약물 목록과 주성분명 표시
    for item in st.session_state.basket:
        st.markdown(f"- **{item}** (성분: *{get_ingredient(item)}*)")
    
    if st.button("💊 통합 DUR 분석 실행", type="primary"):
        st.divider()
        st.subheader("📊 DUR 종합 분석 보고서")
        
        found_any_issue = False
        
        # 1. 병용금기 분석 (조합 분석)
        if len(st.session_state.basket) >= 2 and not dur_db["병용금기"].empty:
            st.markdown("#### ⚠️ 약물 간 상호작용 (병용금기)")
            conf = DUR_CONFIG["병용금기"]
            df_int = dur_db["병용금기"]
            
            for p1, p2 in combinations(st.session_state.basket, 2):
                word1 = get_searchable_word(get_ingredient(p1))
                word2 = get_searchable_word(get_ingredient(p2))
                
                # A-B 또는 B-A 방향 모두 검사
                match = df_int[
                    ((df_int[conf["ing1"]].astype(str).str.lower().str.contains(word1, na=False)) & 
                     (df_int[conf["ing2"]].astype(str).str.lower().str.contains(word2, na=False))) |
                    ((df_int[conf["ing1"]].astype(str).str.lower().str.contains(word2, na=False)) & 
                     (df_int[conf["ing2"]].astype(str).str.lower().str.contains(word1, na=False)))
                ]
                
                if not match.empty:
                    found_any_issue = True
                    msg = match.iloc[0][conf["msg"]] if conf["msg"] in match.columns else "상세정보 없음"
                    st.error(f"**[병용금기 발생] {p1} ↔ {p2}**\n\n사유: {msg}")

        # 2. 개별 약물 주의사항 분석 (연령, 임부, 용량 등)
        st.markdown("#### 📢 개별 약물 주의사항")
        
        for p in st.session_state.basket:
            word = get_searchable_word(get_ingredient(p))
            drug_issues = []
            
            for category in ["연령금기", "임부금기", "효능군중복", "용량주의", "투여기간", "노인주의", "수유부주의"]:
                if dur_db[category].empty: continue
                
                conf = DUR_CONFIG[category]
                match = search_in_df(dur_db[category], conf["ing1"], word)
                
                if not match.empty:
                    found_any_issue = True
                    # 추가 정보(등급, 연령기준 등)가 있으면 함께 표시
                    extra_info = f" ({match.iloc[0][conf['extra']]})" if "extra" in conf and conf["extra"] in match.columns else ""
                    msg = match.iloc[0][conf["msg"]] if conf["msg"] in match.columns else ""
                    drug_issues.append(f"- **[{category}]**{extra_info} : {msg}")
            
            if drug_issues:
                with st.expander(f"💊 {p} 주의사항 확인", expanded=True):
                    for issue in drug_issues:
                        st.warning(issue)
                        
        if not found_any_issue:
            st.success("✅ 현재 약통에 있는 약물 조합 및 개별 약물에 대해 알려진 DUR 주의/금기 사항이 없습니다.")