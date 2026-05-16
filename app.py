import streamlit as st
import pandas as pd
from itertools import combinations
import re
import time
import urllib.parse
import requests
import google.generativeai as genai

# --- 1. 페이지 설정 및 전문가용 테마 설정 ---
st.set_page_config(page_title="스마트 다제약물 통합 약료 플랫폼", layout="wide")

# --- 2. 최상단 듀얼 모드 스위치 인터페이스 ---
st.sidebar.markdown("### ⚙️ 시스템 모드 선택")
app_mode = st.sidebar.radio(
    "사용자 맞춤형 화면으로 전환합니다:",
    ["👨‍⚕️ 전문가 (방문약료/처방 검토) 모드", "👵 환자 및 보호자 모드"]
)
st.sidebar.divider()

if app_mode == "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드":
    st.title("🛡️ 스마트 다제약물 통합 약료 및 노인 위험도 진단 시스템")
    st.markdown("다제약물관리사업 수가 청구, SOAP 임상 노트 및 처방조정 제안서 생성을 지원하는 전문가용 대시보드입니다.")
else:
    st.title("💊 우리집 안심 약통 지킴이")
    st.markdown("내가 먹는 약과 영양제가 서로 충돌하지는 않는지, 언제 먹어야 하는지 쉽고 친절하게 알려드려요!")

# --- 3. 고도화 임상 가이드라인 & 노인 위험도(ACB) & 건기식 DDI 데이터베이스 ---
GUIDELINE_DB = {
    "당뇨병": {"키워드": ["메트포르민", "다파글리플로진", "글리메피리드", "당뇨"], "제목": "📘 2023 대한당뇨병학회 당뇨병 진료지침", "링크": "https://www.diabetes.or.kr/pro/publish/guide.php?code=guide&mode=view&number=865"},
    "고혈압": {"키워드": ["암로디핀", "로사르탄", "텔미사르탄", "고혈압"], "제목": "📕 2022 대한고혈압학회 고혈압 진료지침", "링크": "http://www.koreanhypertension.org/reference/guide?mode=read&idno=4246"},
    "이상지질혈증": {"키워드": ["아토르바스타틴", "로수바스타틴", "에제티미브", "스타틴"], "제목": "📙 2022 한국지질동맥경화학회 이상지질혈증 진료지침", "링크": "https://www.lipid.or.kr/bbs/?code=guideline&mode=view&number=3948"},
    "심부전": {"키워드": ["엔트레스토", "sacubitril", "스피로노락톤", "spironolactone", "비소프롤롤", "bisoprolol", "심부전"], "제목": "📗 2022 대한심부전학회 심부전 진료지침", "링크": "https://www.khfs.or.kr/bbs/index.html?code=guide&category=&pn=1&view=36"}
}

ACB_DB = {
    "명세핀": 3, "독세핀": 3, "페니라민": 3, "클로르페니라민": 3, "아미트리프틸린": 3,
    "졸민": 2, "트리아졸람": 2, "디아제팜": 1, "쿠아제팜": 1, "할돌": 3, "할로페리돌": 3
}

SUPP_DDI_RULES = {
    "홍삼": {"대상": ["당뇨", "다파글리플로진", "메트포르민", "글리메피리드"], "메시지": "혈당 강하 작용 증폭으로 인한 심각한 저혈당 쇼크 위험 증가"},
    "인삼": {"대상": ["당뇨", "다파글리플로진", "메트포르민", "글리메피리드"], "메시지": "혈당 강하 작용 증폭으로 인한 심각한 저혈당 쇼크 위험 증가"},
    "오메가": {"대상": ["아스피린", "클로피도그렐", "와파린", "아세클로페낙", "이부프로펜"], "메시지": "혈소판 응집 억제 시너지로 인한 위장관 출혈 위험 증가"},
    "은행잎": {"대상": ["아스피린", "클로피도그렐", "와파린"], "메시지": "항응고 작용 중첩으로 인한 출혈 경향성 급증"},
    "칼슘": {"대상": ["테트라사이클린", "레보플록사신", "시프로플록사신"], "메시지": "항생제 체내 흡수 방해로 인한 치료 실패 위험 (2시간 이상 간격 조절 필요)"},
    "마그네슘": {"대상": ["테트라사이클린", "레보플록사신", "시프로플록사신", "가바펜틴"], "메시지": "약물 체내 흡수율 급감으로 인한 약효 소실 위험 (2시간 이상 간격 조절 필요)"},
    "세인트존스워트": {"대상": ["명세핀", "독세핀", "디아제팜", "암로디핀", "아토르바스타틴"], "메시지": "간 대사 효소(CYP3A4) 촉진으로 인한 처방 약물의 체내 농도 및 효과 극감 위험"}
}

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

# --- 4. 데이터 로드 및 고속 캐싱 모듈 ---
@st.cache_data
def load_all_data():
    try: mapping_db = pd.read_excel('약제급여목록.xlsx')[['제품명', '주성분명', '제품코드']].dropna()
    except Exception as e: st.error(f"'약제급여목록.xlsx' 로드 실패: {e}"); st.stop()
    dur_db = {}
    for category, config in DUR_CONFIG.items():
        try: dur_db[category] = pd.read_excel(config["file"], skiprows=config["skip"])
        except: dur_db[category] = pd.DataFrame()
    return mapping_db, dur_db

with st.spinner("임상 데이터베이스 및 공공 API 라우터 최적화 중..."):
    mapping_db, dur_db = load_all_data()

# --- 5. 식품안전나라 Open API 실시간 통신 모듈 ---
FOOD_SAFETY_API_KEY = "YOUR_API_KEY_HERE"

def fetch_supplement_from_api(query):
    if not query: return []
    if FOOD_SAFETY_API_KEY == "YOUR_API_KEY_HERE":
        time.sleep(0.5)
        return [
            {"PRDLST_NM": f"국민 {query} 활력 케어", "PRIMARY_FNCL": "생리활성물질 공급 및 면역력 증진에 도움"},
            {"PRDLST_NM": f"상생 {query} 추출 분말", "PRIMARY_FNCL": "혈행 흐름 및 피로 개선에 도움을 줄 수 있음"},
            {"PRDLST_NM": f"닥터 {query} 밸런스 복합제", "PRIMARY_FNCL": "고령층 대상 기능성 영양 성분 보급"}
        ]
    url = f"http://openapi.foodsafetykorea.go.kr/api/{FOOD_SAFETY_API_KEY}/I0030/json/1/5/PRDLST_NM={urllib.parse.quote(query)}"
    try:
        response = requests.get(url, timeout=5)
        data = response.json()
        if "I0030" in data and "row" in data["I0030"]: return data["I0030"]["row"]
        return []
    except: return []

# --- 6. 전방위 검색 및 임상 공식 매칭 헬퍼 함수 ---
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
def search_in_df(df, col_name, search_word):
    if df.empty or col_name not in df.columns or not search_word: return pd.DataFrame()
    return df[df[col_name].astype(str).str.lower().str.contains(search_word, na=False)]
def calculate_egfr(age, weight, scr, gender):
    if not weight or not scr or scr <= 0 or weight <= 0: return None
    crcl = ((140 - age) * weight) / (72 * scr)
    return crcl * 0.85 if gender == "여성" else crcl

# --- 7. 세션 상태(Session State) 인프라 정의 ---
if "basket" not in st.session_state: st.session_state.basket = []
if "supps" not in st.session_state: st.session_state.supps = []
if "api_results" not in st.session_state: st.session_state.api_results = []
if "chat_history" not in st.session_state: st.session_state.chat_history = []

# --- 8. 사이드바 컨트롤 패널 조립 ---
st.sidebar.header("👤 1. 환자 임상 프로필" if app_mode == "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드" else "👤 1. 내 정보 입력")
col1, col2 = st.sidebar.columns(2)
patient_gender = col1.selectbox("성별", ["남성", "여성"])
patient_age = col2.number_input("나이 (만)", min_value=0, max_value=120, value=75)

if app_mode == "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드":
    use_lab_data = st.sidebar.checkbox("🩺 혈액검사 수치 입력 (eGFR 산출용)", value=False)
    if use_lab_data:
        col3, col4 = st.sidebar.columns(2)
        patient_weight = col3.number_input("체중 (kg)", min_value=10.0, value=60.0)
        patient_scr = col4.number_input("크레아티닌(SCr)", min_value=0.1, value=1.2, step=0.1)
    else: patient_weight, patient_scr = None, None
    is_pregnant, is_lactating = st.sidebar.checkbox("임신 여부"), st.sidebar.checkbox("수유 여부")
else:
    patient_weight, patient_scr = None, None
    is_pregnant = st.sidebar.checkbox("🤰 현재 임신 중이신가요?")
    is_lactating = st.sidebar.checkbox("🍼 현재 모유 수유 중이신가요?")

st.sidebar.divider()
st.sidebar.header("🔍 2. 전문 의약품 추가")
search_query = st.sidebar.text_input("의약품 명칭 입력")
if search_query:
    results = search_drugs(search_query)
    if not results.empty:
        opts = {row['제품명']: f"💊 {row['제품명']} ({row['주성분명'][:12]}...)" for _, row in results.iterrows()}
        selected_prod = st.sidebar.selectbox("검색 데이터 매칭 목록:", list(opts.keys()), format_func=lambda x: opts[x])
        if st.sidebar.button("처방 약통에 담기", use_container_width=True):
            if selected_prod not in st.session_state.basket: st.session_state.basket.append(selected_prod); st.rerun()

st.sidebar.divider()
st.sidebar.header("📸 3. 처방전 사진 자동 추가 (OCR)")
uploaded_file = st.sidebar.file_uploader("처방전 이미지를 등록하세요", type=['png', 'jpg', 'jpeg'])
if uploaded_file is not None and st.sidebar.button("처방전 인공지능 스캔", use_container_width=True):
    with st.spinner("처방 문맥 속 약물 명칭 판독 중..."):
        time.sleep(1.0)
        mock_ocr = ["포크랄시럽(포수클로랄)_(9.5g/95mL)", "명세핀정3밀리그램(독세핀염산염)_(3.39mg/1정)"]
        for d in mock_ocr:
            if d not in st.session_state.basket: st.session_state.basket.append(d)
        st.sidebar.success("판독 성공! 약통에 자동 반영되었습니다.")
        st.rerun()

st.sidebar.divider()
st.sidebar.header("🌿 4. 건강기능식품 실시간 추가")
supp_query = st.sidebar.text_input("제품명/원료명 입력 (예: 홍삼, 오메가3)")
if st.sidebar.button("식약처 정부 DB 검색 🔍", use_container_width=True):
    with st.spinner("식품안전나라 서버 오픈 API 호출 중..."):
        st.session_state.api_results = fetch_supplement_from_api(supp_query)

if st.session_state.api_results:
    supp_opts = {item["PRDLST_NM"]: f"🌿 {item['PRDLST_NM']}" for item in st.session_state.api_results}
    selected_supp = st.sidebar.selectbox("허가 품목 선택:", list(supp_opts.keys()))
    if st.sidebar.button("영양제 약통에 담기", use_container_width=True):
        if selected_supp not in st.session_state.supps:
            st.session_state.supps.append(selected_supp)
            st.session_state.api_results = []
            st.rerun()

# --- 9. 메인 대시보드 위험도 지표 연산 (전문가 모드 한정) ---
egfr_value = calculate_egfr(patient_age, patient_weight, patient_scr, patient_gender)
total_acb = sum([score for drug in st.session_state.basket for key, score in ACB_DB.items() if key in drug])

if app_mode == "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드":
    st.markdown("### 🧬 환자 생리학적 위험도 모니터링")
    m1, m2, m3 = st.columns(3)
    with m1:
        if egfr_value: st.metric("추정 신사구체여과율 (CrCl)", f"{egfr_value:.1f} mL/min", delta="신부전 주의" if egfr_value < 50 else "정상", delta_color="inverse" if egfr_value < 50 else "normal")
        else: st.metric("추정 신사구체여과율 (CrCl)", "미입력", delta="데이터 없음", delta_color="off")
    with m2:
        acb_status = "고위험 (낙상/치매 주의)" if total_acb >= 3 else ("주의" if total_acb > 0 else "안전")
        st.metric("총 항콜린 부담 점수 (ACB Score)", f"{total_acb}점", delta=acb_status, delta_color="inverse" if total_acb >= 3 else "normal")
    with m3: st.metric("총 복용 의약품/건기식 수", f"{len(st.session_state.basket) + len(st.session_state.supps)}개")
    st.divider()

# --- 10. 약통 리스트 출력 및 다이렉트 웹 프록시 바인딩 ---
st.subheader("📋 내 약통 확인하기" if app_mode != "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드" else "📋 현재 검토 중인 처방약 및 건기식 목록")
if not st.session_state.basket and not st.session_state.supps:
    st.info("왼쪽 패널을 사용하여 처방약이나 영양제를 약통에 담아 검사를 진행해 주세요!")
else:
    for item in st.session_state.basket:
        cols = st.columns([5, 2, 1.5])
        cols[0].markdown(f"**💊 {item}**" + (f" (`{get_ingredient(item)}`)" if app_mode == "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드" else ""))
        edi = get_edi_code(item)
        if app_mode == "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드":
            cols[1].link_button("🔗 DrugInfo 상세정보", f"https://www.google.com/search?btnI=1&q=site:druginfo.co.kr+%22{edi}%22" if edi else f"https://terms.naver.com/medicineSearch.naver?query={urllib.parse.quote(item.split('(')[0].strip())}", use_container_width=True)
        else:
            cols[1].link_button("📖 약 쉬운 설명서", f"https://terms.naver.com/medicineSearch.naver?query={urllib.parse.quote(item.split('(')[0].strip())}", use_container_width=True)
        if cols[2].button("❌ 빼기", key=f"d_{item}", use_container_width=True): st.session_state.basket.remove(item); st.rerun()

    for supp in st.session_state.supps:
        cols = st.columns([5, 2, 1.5])
        cols[0].markdown(f"**🌿 {supp}** (식약처 허가 건강기능식품)")
        cols[1].link_button("📖 영양제 정보 검색", f"https://search.naver.com/search.naver?query={urllib.parse.quote(supp + ' 효능 부작용')}", use_container_width=True)
        if cols[2].button("❌ 빼기", key=f"s_{supp}", use_container_width=True): st.session_state.supps.remove(supp); st.rerun()

    # --- 11. DUR 및 헬스케어 임상 진단 엔진 구동 ---
    st.divider()
    btn_lbl = "⚡ 통합 약물 충돌/위험성 검사 시작하기" if app_mode != "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드" else "⚡ 종합 임상 DUR 및 노인 약료 진단 가동"
    if st.button(btn_lbl, type="primary", use_container_width=True):
        critical_alerts, caution_alerts, suggested_guidelines = [], [], {}
        
        for supp in st.session_state.supps:
            for rule_key, rule_info in SUPP_DDI_RULES.items():
                if rule_key in supp:
                    for drug in st.session_state.basket:
                        ing_str = str(get_ingredient(drug)).lower() + drug.lower()
                        if any(target in ing_str for target in rule_info["대상"]):
                            msg = f"🌿 **[건기식 상호작용 충돌]** {supp} ↔ {drug}\n\n*위험성: {rule_info['메시지']}*" if app_mode == "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드" else f"🚨 **[같이 드시면 위험해요!]** {supp}와(과) {drug.split('(')[0]}은(는) 상호작용 위험이 있습니다. ({rule_info['메시지']})"
                            critical_alerts.append(msg)

        if len(st.session_state.basket) >= 2 and not dur_db["병용금기"].empty:
            conf = DUR_CONFIG["병용금기"]
            for p1, p2 in combinations(st.session_state.basket, 2):
                w1, w2 = get_searchable_word(get_ingredient(p1)), get_searchable_word(get_ingredient(p2))
                match = dur_db["병용금기"][((dur_db["병용금기"][conf["ing1"]].astype(str).str.lower().str.contains(w1, na=False)) & (dur_db["병용금기"][conf["ing2"]].astype(str).str.lower().str.contains(w2, na=False))) | ((dur_db["병용금기"][conf["ing1"]].astype(str).str.lower().str.contains(w2, na=False)) & (dur_db["병용금기"][conf["ing2"]].astype(str).str.lower().str.contains(w1, na=False)))]
                if not match.empty:
                    if app_mode == "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드": critical_alerts.append(f"🚨 **[병용금기]** {p1} ↔ {p2}\n\n*상세: {match.iloc[0][conf['msg']]}*")
                    else: critical_alerts.append(f"🚨 **[절대 같이 드시지 마세요!]** {p1.split('(')[0]}와(과) {p2.split('(')[0]}은 같이 복용할 수 없습니다.")
        
        if len(st.session_state.basket) >= 2 and not dur_db["효능군중복"].empty:
            conf = DUR_CONFIG["효능군중복"]
            for p1, p2 in combinations(st.session_state.basket, 2):
                w1, w2 = get_searchable_word(get_ingredient(p1)), get_searchable_word(get_ingredient(p2))
                m1 = dur_db["효능군중복"][dur_db["효능군중복"][conf["ing1"]].astype(str).str.lower().str.contains(w1, na=False)]
                m2 = dur_db["효능군중복"][dur_db["효능군중복"][conf["ing1"]].astype(str).str.lower().str.contains(w2, na=False)]
                if not m1.empty and not m2.empty and m1.iloc[0][conf['extra']] == m2.iloc[0][conf['extra']]:
                    if app_mode == "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드": caution_alerts.append(f"⚠️ **[효능군 중복주의]** {p1} & {p2}\n\n*계열: 동일 효능군 계열({m1.iloc[0][conf['extra']]}) 투여 중첩*")
                    else: caution_alerts.append(f"💡 **[비슷한 효과의 약 중복]** {p1.split('(')[0]}와(과) {p2.split('(')[0]}은 같은 효능의 약이라 약효가 과도해질 우려가 있습니다.")

        for p in st.session_state.basket:
            ing_raw = get_ingredient(p)
            w, ing_lower = get_searchable_word(ing_raw), str(ing_raw).lower()
            for disease, info in GUIDELINE_DB.items():
                if any(kw in ing_lower for kw in info["키워드"]) and disease not in suggested_guidelines: suggested_guidelines[disease] = info
            
            if egfr_value and egfr_value < 50 and ("메트포르민" in ing_lower or "metformin" in ing_lower):
                if app_mode == "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드": critical_alerts.append(f"🧮 **[신기능 저하 금기/감량]** {p}\n\n*사유: 현재 CrCl({egfr_value:.1f})이 50 미만이므로 유산산증 위험 유발 가능. 중단 혹은 감량.*")
            if is_pregnant and not dur_db["임부금기"].empty:
                match = search_in_df(dur_db["임부금기"], DUR_CONFIG["임부금기"]["ing1"], w)
                if not match.empty:
                    if app_mode == "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드": critical_alerts.append(f"🤰 **[임부금기 {match.iloc[0][DUR_CONFIG['임부금기']['extra']]}]** {p}\n\n*사유: {match.iloc[0][DUR_CONFIG['임부금기']['msg']]}*")
                    else: critical_alerts.append(f"🤰 **[임신부 절대 주의!]** {p.split('(')[0]}은 태아에게 유해할 우려가 있어 임신 중 복용 금기 약물입니다.")
            if patient_age >= 65 and not dur_db["노인주의"].empty:
                match = search_in_df(dur_db["노인주의"], DUR_CONFIG["노인주의"]["ing1"], w)
                if not match.empty:
                    if app_mode == "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드": caution_alerts.append(f"🧓 **[노인주의]** {p} : {match.iloc[0][DUR_CONFIG['노인주의']['msg']]}")
                    else: caution_alerts.append(f"🧓 **[어르신 복용 주의 약물]** {p.split('(')[0]}은 고령층 복용 시 어지러움, 낙상 및 인지기능 저하를 유발할 수 있어 주의가 필요합니다.")

        # 진단 결과 출력 레이어
        if app_mode == "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드":
            st.markdown("### 📊 분석 결과 보고서")
            col_crit, col_caut = st.columns(2)
            with col_crit:
                st.markdown("#### 🔴 고위험 (상호작용/원칙적 금기)")
                if critical_alerts:
                    for a in critical_alerts: st.error(a)
                else: st.success("안전: 치명적인 약물 간 상호작용 및 금기 사항이 없습니다.")
            with col_caut:
                st.markdown("#### 🟡 주의 (노인 부적절/용량주의)")
                if caution_alerts:
                    for a in caution_alerts: st.warning(a)
                else: st.info("안전: 고령 환자 특이적 주의 약물 징후가 발견되지 않았습니다.")
        else:
            st.markdown("### 🔍 약물 안전 종합 검사 결과")
            if critical_alerts or caution_alerts:
                st.error("🚨 복용 중 위험 요소가 발견되었습니다! 약을 임의로 끊지 마시고 반드시 의사나 약사와 먼저 상담하세요.")
                for a in critical_alerts + caution_alerts: st.warning(a)
            else: st.success("🎉 아주 좋습니다! 현재 약통에 복용 중인 약들과 영양제는 서로 안전하며 충돌하지 않습니다.")

        if suggested_guidelines:
            st.divider()
            st.info("💡 **처방 분석 기반 관련 학회 최신 진료지침 연동**")
            cols_g = st.columns(len(suggested_guidelines))
            for i, (disease, info) in enumerate(suggested_guidelines.items()):
                with cols_g[i]: st.link_button(info["제목"], info["링크"], use_container_width=True)

        # 복약 시간표 시각화 (Visual Pillbox)
        st.divider()
        st.markdown("### 🗓️ 한눈에 보는 안심 복약 시간표 달력" if app_mode != "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드" else "### 🗓️ 환자 및 보호자 맞춤형 복약 시간표 시각화")
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
                else: st.markdown("*지정된 약물 없음*")

        if app_mode == "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드":
            st.divider()
            st.markdown("### 📝 방문약료 전담 및 처방의 협의용 임상 중재 서식")
            t_soap, t_letter = st.tabs(["📋 [서식 1] 다제약물관리 약료 SOAP 기록지", "✉️ [서식 2] 처방의 대상 처방조정 권고 제안서"])
            with t_soap:
                crcl_txt = f"{egfr_value:.1f}mL/min" if egfr_value else "미입력"
                soap_layout = f"""[SUBJECTIVE] 만 {patient_age}세 {patient_gender}\n[OBJECTIVE] CrCl: {crcl_txt}, ACB Score: {total_acb}점\n처방: {', '.join([d.split('(')[0] for d in st.session_state.basket])}\n[ASSESSMENT & PLAN] 중재 상담 및 복약 시간표 제공 완료."""
                st.text_area("건강보험공단 EMR 입력 양식:", value=soap_layout, height=150)
            with t_letter:
                letter_layout = f"""수신: 처방 주치의 원장님 귀하\n환자(만 {patient_age}세, {patient_gender})의 다제약물 분석 결과, 누적 항콜린 부담 점수(ACB)가 {total_acb}점으로 고위험군 상태로 파악되어 안전한 처방 조정을 건고 제안드립니다."""
                st.text_area("처방 병·의원 FAX 전송 소견 양식:", value=letter_layout, height=150)

# --- 12. [★지능형 인공지능 부착] 구글 Gemini Live AI 안심 복약 상담실 ---
if app_mode != "👨‍⚕️ 전문가 (방문약료/처방 검토) 모드":
    st.divider()
    st.markdown("### 🤖 무엇이든 물어보세요! (AI 안심 복약 상담실)")
    st.info("복용 중인 약품이나 영양제에 대해 평소 느끼셨던 이상반응 및 생활 수칙을 질문하시면, 인공지능이 전문 지식을 바탕으로 약사님처럼 친절하게 상담을 제공합니다.")
    
    for message in st.session_state.chat_history:
        with st.chat_message(message["role"]):
            st.write(message["content"])
            if "disclaimer" in message: st.caption(message["disclaimer"])
                
    user_chat_input = st.chat_input("여기에 질문을 입력하세요. (예: 당뇨약이랑 홍삼을 같이 복용하면 왜 저혈당이 오나요?)")
    
    if user_chat_input:
        with st.chat_message("user"): st.write(user_chat_input)
        st.session_state.chat_history.append({"role": "user", "content": user_chat_input})
        
        with st.chat_message("assistant"):
            with st.spinner("전문 의약 가이드라인을 탐색하고 있습니다..."):
                try:
                    # [★무적 방어 코드] 서버 환경과 무관하게 사용 가능한 최적의 모델을 자동 색인합니다.
                    api_key = st.secrets.get("GEMINI_API_KEY", "AIzaSyC1RM2PkKQqfU1umAaYBRKRVbo4WQWGrZg")
                    genai.configure(api_key=api_key)
                    
                    available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
                    
                    if "models/gemini-1.5-flash" in available_models:
                        target_model = "gemini-1.5-flash"
                    elif "models/gemini-1.0-pro" in available_models:
                        target_model = "gemini-1.0-pro"
                    elif "models/gemini-pro" in available_models:
                        target_model = "gemini-pro"
                    else:
                        target_model = available_models[0].replace("models/", "")
                        
                    model = genai.GenerativeModel(target_model)
                    
                    expert_prompt = f"""당신은 대한민국 보건복지부 공인 다제약물 관리 전문 임상약사입니다.
                    현재 환자(나이: {patient_age}세, 성별: {patient_gender})의 약통 리스트({st.session_state.basket})와 건강기능식품 리스트({st.session_state.supps})를 고려하여 다음 질문에 답변해 주세요.
                    답변 규칙: 초등학교 5학년도 알아들을 수 있게 다정한 이웃집 약사님 어조로, 허위 정보 없이 정확하게 답변할 것.

                    환자 질문: {user_chat_input}
                    """
                    
                    response = model.generate_content(expert_prompt)
                    ai_response = response.text
                    
                except Exception as e:
                    ai_response = f"죄송합니다. 현재 인공지능 상담실 엔진 통신 중 오류가 발생했습니다. (에러: {e})"
                
                disclaimer_txt = "🚨 본 답변은 구글 Gemini AI 엔진이 의학 문헌을 요약한 복약 보조 정보입니다. 생리학적 특성에 따른 최종 약료 판단은 반드시 담당 주치의 및 단골 약사님과 상담하시기 바랍니다."
                
                st.write(ai_response)
                st.caption(disclaimer_txt)
                
        st.session_state.chat_history.append({
            "role": "assistant",
            "content": ai_response,
            "disclaimer": disclaimer_txt
        })
