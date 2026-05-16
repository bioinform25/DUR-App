import streamlit as st
import pandas as pd
from itertools import combinations
import re
import time
import urllib.parse

# --- 1. 페이지 설정 및 전문가용 테마 설정 ---
st.set_page_config(page_title="스마트 통합 DUR 및 노인 약료 진단 시스템", layout="wide")
st.title("🛡️ 스마트 다제약물 통합 약료 및 노인 위험도 진단 시스템")
st.markdown("방문약료 및 임상 약료 전문가를 위한 K-노인약료 특화 다제약물 제어 및 SOAP 자동 생성 솔루션입니다.")

# --- 2. 최신 임상 진료 지침(Guideline) 링크 사전 정의 ---
GUIDELINE_DB = {
    "당뇨병": {"키워드": ["메트포르민", "다파글리플로진", "글리메피리드", "당뇨"], "제목": "📘 2023 당뇨병 진료지침", "링크": "https://www.diabetes.or.kr/pro/publish/guide.php?code=guide&mode=view&number=865"},
    "고혈압": {"키워드": ["암로디핀", "로사르탄", "텔미사르탄", "고혈압"], "제목": "📕 2022 고혈압 진료지침", "링크": "http://www.koreanhypertension.org/reference/guide?mode=read&idno=4246"},
    "이상지질혈증": {"키워드": ["아토르바스타틴", "로수바스타틴", "에제티미브", "스타틴"], "제목": "📙 2022 이상지질혈증 진료지침", "링크": "https://www.lipid.or.kr/bbs/?code=guideline&mode=view&number=3948"},
    "심부전": {"키워드": ["엔트레스토", "sacubitril", "스피로노락톤", "spironolactone", "비소프롤롤", "bisoprolol", "심부전"], "제목": "📗 2022 대한심부전학회 심부전 진료지침", "링크": "https://www.khfs.or.kr/bbs/index.html?code=guide&category=&pn=1&view=36"}
}

# 노인 항콜린 부담 점수 (ACB Score) DB
ACB_DB = {
    "명세핀": 3, "독세핀": 3, "페니라민": 3, "클로르페니라민": 3, "아미트리프틸린": 3,
    "졸민": 2, "트리아졸람": 2, "디아제팜": 1, "쿠아제팜": 1, "할돌": 3, "할로페리돌": 3
}

# 한국형 건강기능식품 DDI 매핑 DB
SUPP_DDI_DB = {
    "홍삼/인삼": {"대상": ["당뇨", "다파글리플로진", "메트포르민", "글리메피리드"], "메시지": "혈당 강하 작용 증폭으로 인한 심각한 저혈당 쇼크 위험 증가"},
    "오메가3/크릴오일": {"대상": ["아스피린", "클로피도그렐", "와파린", "아세클로페낙", "이부프로펜"], "메시지": "혈소판 응집 억제 시너지로 인한 위장관 출혈 위험 증가"},
    "은행잎추출물(징코)": {"대상": ["아스피린", "클로피도그렐", "와파린"], "메시지": "항응고 작용 중첩으로 인한 출혈 경향성 급증"}
}

# DUR 데이터셋 파일 구조 정의 지도
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

# --- 3. 데이터 로드 및 고속 캐싱 ---
@st.cache_data
def load_all_data():
    try: mapping_db = pd.read_excel('약제급여목록.xlsx')[['제품명', '주성분명', '제품코드']].dropna()
    except Exception as e: st.error(f"'약제급여목록.xlsx' 로드 실패: {e}"); st.stop()
    dur_db = {}
    for category, config in DUR_CONFIG.items():
        try: dur_db[category] = pd.read_excel(config["file"], skiprows=config["skip"])
        except: dur_db[category] = pd.DataFrame()
    return mapping_db, dur_db

with st.spinner("임상 데이터베이스 및 노인 약료 알고리즘 최적화 중..."):
    mapping_db, dur_db = load_all_data()

# --- 4. 완벽하게 보완된 헬퍼 함수 ---
def search_drugs(query):
    if not query: return pd.DataFrame()
    q = query.strip().lower()
    return mapping_db[mapping_db['제품명'].str.contains(q, na=False, case=False) | mapping_db['주성분명'].str.contains(q, na=False, case=False)]

def get_ingredient(product_name):
    match = mapping_db[mapping_db['제품명'] == product_name]
    return match.iloc[0]['주성분명'] if not match.empty else product_name

def get_edi_code(product_name):
    match = mapping_db[mapping_db['제품명'] == product_name]
    return str(match.iloc[0]['제품코드']).split('.')[0].strip() if not match.empty else ""

def get_searchable_word(ing_string):
    if not ing_string: return ""
    clean = re.sub(r'[^a-zA-Z\s]', ' ', str(ing_string)).strip().lower()
    return clean.split()[0] if clean else str(ing_string)

# [★버그 해결] 누락되었던 데이터프레임 내 성분명 정밀 검색 함수 탑재
def search_in_df(df, col_name, search_word):
    if df.empty or col_name not in df.columns or not search_word:
        return pd.DataFrame()
    return df[df[col_name].astype(str).str.lower().str.contains(search_word, na=False)]

def calculate_egfr(age, weight, scr, gender):
    if not weight or not scr or scr <= 0 or weight <= 0: return None
    crcl = ((140 - age) * weight) / (72 * scr)
    return crcl * 0.85 if gender == "여성" else crcl

if "basket" not in st.session_state: st.session_state.basket = []
if "supps" not in st.session_state: st.session_state.supps = []

# --- 5. 사이드바 구성 ---
st.sidebar.header("👤 1. 환자 임상 프로필")
col1, col2 = st.sidebar.columns(2)
patient_gender = col1.selectbox("성별", ["남성", "여성"])
patient_age = col2.number_input("나이 (만)", min_value=0, max_value=120, value=75)

use_lab_data = st.sidebar.checkbox("스마트 진단: 혈액검사 수치 입력 (선택사항)", value=False)
if use_lab_data:
    col3, col4 = st.sidebar.columns(2)
    patient_weight = col3.number_input("체중 (kg)", min_value=10.0, value=60.0)
    patient_scr = col4.number_input("혈청 크레아티닌(SCr)", min_value=0.1, value=1.2, step=0.1)
else:
    patient_weight = None
    patient_scr = None

is_pregnant = st.sidebar.checkbox("임신 여부 (임부필터 활성)")
is_lactating = st.sidebar.checkbox("수유 여부 (수유부필터 활성)")

st.sidebar.divider()
st.sidebar.header("🔍 2. 전문 의약품 추가")
search_query = st.sidebar.text_input("상품명 또는 영문 성분명 입력")
if search_query:
    results = search_drugs(search_query)
    if not results.empty:
        options_map = {row['제품명']: f"💊 {row['제품명']} ({row['주성분명'][:15]}...)" for _, row in results.iterrows()}
        selected_prod = st.sidebar.selectbox("결과 선택:", options=list(options_map.keys()), format_func=lambda x: options_map[x])
        if st.sidebar.button("처방 약통에 추가", use_container_width=True):
            if selected_prod not in st.session_state.basket: st.session_state.basket.append(selected_prod); st.rerun()

st.sidebar.divider()
st.sidebar.header("🌿 3. 복용 중인 건강기능식품 추가")
selected_supp = st.sidebar.selectbox("한국 다빈도 건기식", ["선택 안 함"] + list(SUPP_DDI_DB.keys()))
if selected_supp != "선택 안 함" and st.sidebar.button("건기식 약통에 추가", use_container_width=True):
    if selected_supp not in st.session_state.supps: st.session_state.supps.append(selected_supp); st.rerun()

# --- 6. 메인 화면 상단 위험도 대시보드 ---
egfr_value = calculate_egfr(patient_age, patient_weight, patient_scr, patient_gender)
total_acb = sum([score for drug in st.session_state.basket for key, score in ACB_DB.items() if key in drug])

st.markdown("### 🧬 환자 생리학적 위험도 모니터링")
m1, m2, m3 = st.columns(3)
with m1:
    if egfr_value:
        st.metric(label="추정 신사구체여과율 (CrCl)", value=f"{egfr_value:.1f} mL/min", delta="신부전 주의" if egfr_value < 50 else "정상", delta_color="inverse" if egfr_value < 50 else "normal")
    else:
        st.metric(label="추정 신사구체여과율 (CrCl)", value="미입력", delta="데이터 없음", delta_color="off")
with m2:
    acb_status = "고위험 (낙상/치매 주의)" if total_acb >= 3 else ("주의" if total_acb > 0 else "안전")
    st.metric(label="총 항콜린 부담 점수 (ACB Score)", value=f"{total_acb}점", delta=acb_status, delta_color="inverse" if total_acb >= 3 else "normal")
with m3:
    st.metric(label="총 복용 약물 수", value=f"{len(st.session_state.basket) + len(st.session_state.supps)}개", delta="다제약물(Polypharmacy)" if len(st.session_state.basket)>=5 else "적정", delta_color="inverse" if len(st.session_state.basket)>=5 else "normal")

st.divider()

# --- 7. 내 약통 목록 리스트 ---
st.subheader("📋 현재 복용 중인 처방약 및 건기식")
if not st.session_state.basket and not st.session_state.supps:
    st.info("왼쪽 사이드바에서 환자 정보를 입력하고 약물을 추가해 주세요.")
else:
    for item in st.session_state.basket:
        c1, c2, c3 = st.columns([5, 2, 1.5])
        c1.markdown(f"**💊 {item}** (`{get_ingredient(item)}`)")
        edi = get_edi_code(item)
        c2.link_button("🔗 DrugInfo 상세정보", f"https://www.google.com/search?btnI=1&q=site:druginfo.co.kr+%22{edi}%22" if edi else f"https://terms.naver.com/medicineSearch.naver?query={urllib.parse.quote(item.split('(')[0].strip())}", use_container_width=True)
        if c3.button("❌ 삭제", key=f"d_{item}", use_container_width=True): st.session_state.basket.remove(item); st.rerun()

    for supp in st.session_state.supps:
        c1, c2, c3 = st.columns([5, 2, 1.5])
        c1.markdown(f"**🌿 {supp}** (건강기능식품/한약재)")
        c2.markdown("") 
        if c3.button("❌ 삭제", key=f"s_{supp}", use_container_width=True): st.session_state.supps.remove(supp); st.rerun()

    # --- 8. 고도화된 통합 진단 분석 엔진 가동 ---
    st.divider()
    if st.button("⚡ 종합 임상 DUR 및 노인 약료 진단 가동", type="primary", use_container_width=True):
        critical_alerts, caution_alerts, suggested_guidelines = [], [], {}
        
        # 건기식 상호작용 검사
        for supp in st.session_state.supps:
            supp_info = SUPP_DDI_DB[supp]
            for drug in st.session_state.basket:
                ing_str = str(get_ingredient(drug)).lower() + drug.lower()
                if any(target in ing_str for target in supp_info["대상"]):
                    critical_alerts.append(f"🌿 **[건기식 상호작용 충돌]** {supp} ↔ {drug}\n\n*위험성: {supp_info['메시지']}*")

        # 8대 DUR 다각도 복합 분석 루프
        if len(st.session_state.basket) >= 2:
            if not dur_db["병용금기"].empty:
                conf = DUR_CONFIG["병용금기"]
                for p1, p2 in combinations(st.session_state.basket, 2):
                    w1, w2 = get_searchable_word(get_ingredient(p1)), get_searchable_word(get_ingredient(p2))
                    match = dur_db["병용금기"][((dur_db["병용금기"][conf["ing1"]].astype(str).str.lower().str.contains(w1, na=False)) & (dur_db["병용금기"][conf["ing2"]].astype(str).str.lower().str.contains(w2, na=False))) | ((dur_db["병용금기"][conf["ing1"]].astype(str).str.lower().str.contains(w2, na=False)) & (dur_db["병용금기"][conf["ing2"]].astype(str).str.lower().str.contains(w1, na=False)))]
                    if not match.empty: critical_alerts.append(f"🚨 **[병용금기]** {p1} ↔ {p2}\n\n*사유: {match.iloc[0][conf['msg']]}*")
            
            if not dur_db["효능군중복"].empty:
                conf = DUR_CONFIG["효능군중복"]
                for p1, p2 in combinations(st.session_state.basket, 2):
                    w1, w2 = get_searchable_word(get_ingredient(p1)), get_searchable_word(get_ingredient(p2))
                    m1 = dur_db["효능군중복"][dur_db["효능군중복"][conf["ing1"]].astype(str).str.lower().str.contains(w1, na=False)]
                    m2 = dur_db["효능군중복"][dur_db["효능군중복"][conf["ing1"]].astype(str).str.lower().str.contains(w2, na=False)]
                    if not m1.empty and not m2.empty and m1.iloc[0][conf['extra']] == m2.iloc[0][conf['extra']]:
                        caution_alerts.append(f"⚠️ **[효능군 중복주의]** {p1} & {p2}\n\n*사유: 동일 계열 ({m1.iloc[0][conf['extra']]}) 중복 투여*")
        
        # 개별 약물 분석 루프 (임부, 노인, 수유부, 연령, 용량, 투여기간 전체 반영)
        for p in st.session_state.basket:
            ing_raw = get_ingredient(p)
            w, ing_lower = get_searchable_word(ing_raw), str(ing_raw).lower()
            
            for disease, info in GUIDELINE_DB.items():
                if any(kw in ing_lower for kw in info["키워드"]) and disease not in suggested_guidelines: suggested_guidelines[disease] = info
            
            if egfr_value and egfr_value < 50 and ("메트포르민" in ing_lower or "metformin" in ing_lower):
                critical_alerts.append(f"🧮 **[신기능 저하 금기/감량]** {p}\n\n*사유: 현재 CrCl({egfr_value:.1f})이 50 미만이므로 유산산증 위험 급증. 감량 또는 중단 고려.*")

            if is_pregnant and not dur_db["임부금기"].empty:
                conf = DUR_CONFIG["임부금기"]
                match = search_in_df(dur_db["임부금기"], conf["ing1"], w)
                if not match.empty: critical_alerts.append(f"🤰 **[임부금기 {match.iloc[0][conf['extra']]}]** {p}\n\n*사유: {match.iloc[0][conf['msg']]}*")

            if is_lactating and not dur_db["수유부주의"].empty:
                conf = DUR_CONFIG["수유부주의"]
                match = search_in_df(dur_db["수유부주의"], conf["ing1"], w)
                if not match.empty: caution_alerts.append(f"🍼 **[수유부주의]** {p} : 모유 이행 및 영아 이상반응 우려 ({match.iloc[0][conf['msg']]})")

            if patient_age >= 65 and not dur_db["노인주의"].empty:
                conf = DUR_CONFIG["노인주의"]
                match = search_in_df(dur_db["노인주의"], conf["ing1"], w)
                if not match.empty: caution_alerts.append(f"🧓 **[노인주의]** {p} : {match.iloc[0][conf['msg']]}")

            for cat in ["연령금기", "용량주의", "투여기간"]:
                if dur_db[cat].empty: continue
                conf = DUR_CONFIG[cat]
                match = search_in_df(dur_db[cat], conf["ing1"], w)
                if not match.empty:
                    extra_txt = f" ({match.iloc[0][conf['extra']]})" if "extra" in conf and str(match.iloc[0][conf['extra']]) != 'nan' else ""
                    caution_alerts.append(f"📢 **[{cat}{extra_txt}]** {p} : {match.iloc[0][conf['msg']]}")

        # --- 진단 대시보드 시각화 출력 ---
        st.markdown("### 📊 분석 결과 보고서")
        col_crit, col_caut = st.columns(2)
        with col_crit:
            st.markdown("#### 🔴 고위험 (상호작용/금기 사항)")
            if critical_alerts: 
                for a in critical_alerts: st.error(a)
            else: st.success("발견된 치명적 금기 사항이 없습니다.")
        with col_caut:
            st.markdown("#### 🟡 주의 (노인/용량/투여기간)")
            if caution_alerts: 
                for a in caution_alerts: st.warning(a)
            else: st.info("발견된 임상 주의 사항이 없습니다.")

        # --- 가이드라인 연동 추천 시스템 ---
        if suggested_guidelines:
            st.divider()
            st.info("💡 **처방 성분 분석 기반 최신 진료가이드라인 연동**")
            cols_g = st.columns(len(suggested_guidelines))
            for i, (disease, info) in enumerate(suggested_guidelines.items()):
                with cols_g[i]: st.link_button(info["제목"], info["링크"], use_container_width=True)

        # --- 스마트 시각적 복약 시간표 (Visual Pillbox) ---
        st.divider()
        st.markdown("### 🗓️ 환자 및 보호자용 맞춤형 복약 달력 (Visual Pillbox)")
        pillbox = {"🌅 아침 (식후)": [], "☀️ 점심 (식후)": [], "🌃 저녁 (식후)": [], "🌙 취침 전": []}
        for drug in st.session_state.basket:
            if any(x in drug for x in ["수면", "졸민", "명세핀", "독세핀", "트리아졸람", "디아제팜"]): pillbox["🌙 취침 전"].append(drug)
            elif any(x in drug for x in ["당뇨", "메트포르민"]): pillbox["🌅 아침 (식후)"].append(drug); pillbox["🌃 저녁 (식후)"].append(drug)
            elif any(x in drug for x in ["혈압", "암로디핀"]): pillbox["🌅 아침 (식후)"].append(drug)
            else: pillbox["🌅 아침 (식후)"].append(drug); pillbox["🌃 저녁 (식후)"].append(drug)
        for supp in st.session_state.supps: pillbox["☀️ 점심 (식후)"].append(f"🌿 {supp}")

        pb_cols = st.columns(4)
        for idx, (time_slot, drugs) in enumerate(pillbox.items()):
            with pb_cols[idx]:
                st.markdown(f"**{time_slot}**")
                if drugs:
                    for d in list(set(drugs)): st.info(d.split('(')[0])
                else: st.markdown("*복용 약 없음*")

        # --- SOAP 임상 노트 리포트 작성 ---
        st.divider()
        st.markdown("### 📝 환자 맞춤형 약료 임상 노트 (SOAP 서식)")
        crcl_text = f"{egfr_value:.1f}mL/min" if egfr_value else "미입력 (Lab 데이터 없음)"
        weight_text = f"체중 {patient_weight}kg" if patient_weight else "체중 미입력"
        
        soap_note = f"""[SUBJECTIVE] 만 {patient_age}세 {patient_gender}, {weight_text}
[OBJECTIVE] CrCl: {crcl_text}, 총 항콜린점수(ACB): {total_acb}점
처방약물: {', '.join([d.split('(')[0] for d in st.session_state.basket])}
복용건기식: {', '.join(st.session_state.supps) if st.session_state.supps else '없음'}

[ASSESSMENT & PLAN]
- 노인 생리학적 위험도 지표 평가: {"낙상 및 인지 기능 저하 고위험 처방군" if total_acb >=3 else "항콜린제 누적 위험 정상"}
- 임상 권고 사항: {"발견된 병용 위험이 있으므로 처방의 협의 및 건기식 제어 복약지도." if critical_alerts else "DUR 적합성 확인됨. 복약 달력 기준으로 순응도 유지 지도."}"""
        st.text_area("방문약료 및 EMR 시스템 서식 복사용:", value=soap_note, height=200)
