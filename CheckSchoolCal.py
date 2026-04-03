import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, timedelta, datetime
from korean_lunar_calendar import KoreanLunarCalendar
import os

# ============================================================
# 페이지 설정
# ============================================================
st.set_page_config(page_title="학사일정 점검 시스템", page_icon="📋", layout="wide")

# ============================================================
# 1. 음력 → 양력 변환
# ============================================================
def lunar_to_solar(lunar_year, lunar_month, lunar_day):
    cal = KoreanLunarCalendar()
    cal.setLunarDate(lunar_year, lunar_month, lunar_day, False)
    return date(*map(int, cal.SolarIsoFormat().split('-')))

# ============================================================
# 2. 대체공휴일 계산
# ============================================================
def next_weekday_not_in(start, existing):
    """start 이후 첫 평일(기존 공휴일과 겹치지 않는)"""
    d = start
    while d.weekday() >= 5 or d in existing:
        d += timedelta(days=1)
    return d

def apply_substitute_single(holiday_date, name, all_dates, no_substitute=False):
    """일반 공휴일(1일짜리)의 대체공휴일 계산"""
    result = [(holiday_date, name)]
    if no_substitute:
        return result

    if holiday_date.weekday() >= 5:  # 토 or 일
        occupied = all_dates | {holiday_date}
        sub = next_weekday_not_in(holiday_date + timedelta(days=1), occupied)
        result.append((sub, f"{name} 대체공휴일"))

    return result

def apply_substitute_holiday3(center_date, name, all_dates):
    """명절(3일 연휴: 설날/추석)의 대체공휴일 계산"""
    day_before = center_date - timedelta(days=1)
    day_after  = center_date + timedelta(days=1)
    base_days  = [day_before, center_date, day_after]

    result = [(d, f"{name} 연휴") for d in base_days]

    if center_date.weekday() >= 5:
        occupied = all_dates | set(base_days)
        sub = next_weekday_not_in(day_after + timedelta(days=1), occupied)
        result.append((sub, f"{name} 대체공휴일"))

    return result

# ============================================================
# 3. 학년도 기준 공휴일 자동 생성 (윤년 유연성 반영)
# ============================================================
def generate_holidays(school_year):
    """학년도(school_year) 기준 공휴일 자동 생성"""
    Y = school_year
    all_dates = set()
    holidays = []

    # --- 음력 공휴일 ---
    buddha         = lunar_to_solar(Y, 4, 8)
    chuseok_center = lunar_to_solar(Y, 8, 15)
    seollal_center = lunar_to_solar(Y + 1, 1, 1)

    # --- 양력 고정 공휴일 ---
    fixed = [
        (date(Y, 3, 1),     "삼일절",     False),
        (date(Y, 5, 5),     "어린이날",   False),
        (date(Y, 6, 6),     "현충일",     True),   # 대체 미적용
        (date(Y, 8, 15),    "광복절",     False),
        (date(Y, 10, 3),    "개천절",     False),
        (date(Y, 10, 9),    "한글날",     False),
        (date(Y, 12, 25),   "기독탄신일", False),
        (date(Y + 1, 1, 1), "신정",       False),
    ]
    if Y >= 2026:
        fixed.append((date(Y, 7, 17), "제헌절", False))

    for hdate, hname, no_sub in fixed:
        for d, n in apply_substitute_single(hdate, hname, all_dates, no_sub):
            all_dates.add(d)
            holidays.append((d, n))

    for d, n in apply_substitute_single(buddha, "부처님오신날", all_dates, False):
        all_dates.add(d)
        holidays.append((d, n))

    for d, n in apply_substitute_holiday3(chuseok_center, "추석", all_dates):
        all_dates.add(d)
        holidays.append((d, n))

    for d, n in apply_substitute_holiday3(seollal_center, "설날", all_dates):
        all_dates.add(d)
        holidays.append((d, n))

    # [보완] 학년도 범위 필터: Y.3.1 ~ (Y+1).2.28 or 29 (윤년 자동 계산)
    start = date(Y, 3, 1)
    end   = date(Y + 1, 3, 1) - timedelta(days=1)
    
    holidays = [(d, n) for d, n in holidays if start <= d <= end]

    # 평일만 필터
    holidays = [(d, n) for d, n in holidays if d.weekday() < 5]

    return dict(sorted(holidays, key=lambda x: x[0]))

# ============================================================
# 4. 유틸리티 함수 (성능 개선)
# ============================================================
def count_weekdays(start_date, end_date):
    """Numpy를 활용한 평일 계산 최적화"""
    if not start_date or not end_date or start_date > end_date:
        return 0
    # np.busday_count는 종료일을 포함하지 않으므로 +1일 처리
    end_inclusive = end_date + timedelta(days=1)
    return int(np.busday_count(start_date, end_inclusive))

def count_holidays_in_range(start_date, end_date, holidays_set):
    return sum(1 for h in holidays_set if start_date <= h <= end_date and h.weekday() < 5)

def to_date(val):
    if val is None: return None
    try:
        if pd.isna(val): return None
    except: pass
    if isinstance(val, (datetime, pd.Timestamp)): return val.date()
    if isinstance(val, date): return val
    if isinstance(val, str):
        val = val.strip()
        if not val or val.lower() == 'nat': return None
        try: return pd.to_datetime(val).date()
        except: return None
    return None

def to_int(val):
    if val is None: return 0
    try:
        if pd.isna(val): return 0
    except: pass
    try: return int(float(val))
    except: return 0

def get_nov_first_week_deadline(year):
    """
    해당 연도의 11월 첫 주 일요일 날짜를 계산합니다.
    (단, 11월 1일이 토/일요일인 경우 실질적인 학사일정 상 첫 주는 다음 주이므로 +7일 처리)
    """
    first_day = date(year, 11, 1)
    if first_day.weekday() >= 5:  # 5: 토요일, 6: 일요일
        days_to_sunday = 6 - first_day.weekday()
        return first_day + timedelta(days=days_to_sunday + 7)
    else:
        days_to_sunday = 6 - first_day.weekday()
        return first_day + timedelta(days=days_to_sunday)

# ============================================================
# 5. 핵심 점검 함수
# ============================================================
def extract_column_indices(df):
    header_area = df.iloc[:10].copy()
    header_area = header_area.ffill(axis=1).ffill(axis=0)
    
    col_map = {"disc_dates_1": [], "disc_dates_2": []}
    
    for col in range(header_area.shape[1]):
        text = "".join(str(x).replace('\n', '').replace(' ', '') for x in header_area.iloc[:, col] if pd.notna(x))
        
        if "연번" in text and "학교명" not in text and "serial" not in col_map:
            col_map["serial"] = col
        elif "학교명" in text and "school_name" not in col_map:
            col_map["school_name"] = col
        elif "학교급" in text and "school_level" not in col_map:
            col_map["school_level"] = col
        elif "적용학년" in text and "grade" not in col_map:
            col_map["grade"] = col
            
        elif "수업일수" in text and "1학기" in text and "재량" not in text and "고사" not in text and "sem1" not in col_map:
            col_map["sem1"] = col
        elif "수업일수" in text and "2학기" in text and "재량" not in text and "고사" not in text and "sem2" not in col_map:
            col_map["sem2"] = col
            
        elif ("개학일" in text or "입학일" in text) and "여름방학" not in text and "겨울방학" not in text and "open" not in col_map:
            col_map["open"] = col
            
        elif "여름방학" in text and "방학식일" in text and "s_close" not in col_map:
            col_map["s_close"] = col
        elif "여름방학" in text and "개학식일" in text and "s_open" not in col_map:
            col_map["s_open"] = col
            
        elif "겨울방학" in text and "방학식일" in text and "w_close" not in col_map:
            col_map["w_close"] = col
        elif "겨울방학" in text and "개학식일" in text and "w_open" not in col_map:
            col_map["w_open"] = col
            
        elif "종업식" in text and "end_12" not in col_map:
            col_map["end_12"] = col
        elif "졸업식" in text and "grad_3" not in col_map:
            col_map["grad_3"] = col
            
        # [핵심 수정] "중간, 기말고사" 병합 셀의 단어 간섭을 막기 위해 검색어 밀착("2학기기말")
        elif "2학기기말" in text and "종료" in text and "term2_final_end" not in col_map:
            col_map["term2_final_end"] = col
            
        elif "재량" in text and "1학기" in text and "수" in text and "날짜" not in text and "disc1" not in col_map:
            col_map["disc1"] = col
        elif "재량" in text and "2학기" in text and "수" in text and "날짜" not in text and "disc2" not in col_map:
            col_map["disc2"] = col
            
        elif "재량" in text and "1학기" in text and "날짜" in text:
            col_map["disc_dates_1"].append(col)
        elif "재량" in text and "2학기" in text and "날짜" in text:
            col_map["disc_dates_2"].append(col)

    defaults = {
        "serial": 0, "school_level": 3, "school_name": 6, "grade": 12, "sem1": 13, "sem2": 14,
        "open": 16, "s_close": 17, "s_open": 18, "w_close": 19, "w_open": 20,
        "end_12": 21, "grad_3": 22, "term2_final_end": 33, "disc1": 51, "disc2": 52
    }
    for k, v in defaults.items():
        if k not in col_map:
            col_map[k] = v
            
    return col_map

def check_school(file_name, df, holidays_dict, school_year):
    school_name = os.path.basename(file_name).split('_')[0]
    errors, details = [], []
    col_idx = extract_column_indices(df)

    rows_12, rows_3 = [], []
    for idx in range(df.shape[0]):
        row = df.iloc[idx]
        grade = str(row.iloc[col_idx["grade"]]).strip() if pd.notna(row.iloc[col_idx["grade"]]) else ''
        serial = str(row.iloc[col_idx["serial"]]).strip() if pd.notna(row.iloc[col_idx["serial"]]) else ''
        
        if serial == '예시': continue
        if grade == '1~2': rows_12.append(row)
        elif grade == '3': rows_3.append(row)

    if not rows_12 or not rows_3:
        errors.append(f"[{school_name}] 데이터 행(1~2학년/3학년)을 찾을 수 없습니다.")
        return errors, details

    for r12, r3 in zip(rows_12, rows_3):
        sname = str(r12.iloc[col_idx["school_name"]]) if pd.notna(r12.iloc[col_idx["school_name"]]) else school_name
        school_level = str(r12.iloc[col_idx["school_level"]]) if pd.notna(r12.iloc[col_idx["school_level"]]) else ''

        open_12      = to_date(r12.iloc[col_idx["open"]])
        open_3       = to_date(r3.iloc[col_idx["open"]])
        s_close_12   = to_date(r12.iloc[col_idx["s_close"]])
        s_open_12    = to_date(r12.iloc[col_idx["s_open"]])
        w_close_12   = to_date(r12.iloc[col_idx["w_close"]])
        w_open_12    = to_date(r12.iloc[col_idx["w_open"]])
        end_12       = to_date(r12.iloc[col_idx["end_12"]])
        s_close_3    = to_date(r3.iloc[col_idx["s_close"]])
        s_open_3     = to_date(r3.iloc[col_idx["s_open"]])
        w_close_3    = to_date(r3.iloc[col_idx["w_close"]])
        w_open_3     = to_date(r3.iloc[col_idx["w_open"]])
        grad_3       = to_date(r3.iloc[col_idx["grad_3"]])
        if grad_3 is None: grad_3 = to_date(r3.iloc[col_idx["end_12"]])
        
        # 3학년 2학기 기말고사 종료일
        term2_final_end_3 = to_date(r3.iloc[col_idx["term2_final_end"]])

        sem1_12, sem2_12   = to_int(r12.iloc[col_idx["sem1"]]), to_int(r12.iloc[col_idx["sem2"]])
        sem1_3,  sem2_3    = to_int(r3.iloc[col_idx["sem1"]]),  to_int(r3.iloc[col_idx["sem2"]])
        disc1_12, disc2_12 = to_int(r12.iloc[col_idx["disc1"]]), to_int(r12.iloc[col_idx["disc2"]])
        disc1_3,  disc2_3  = to_int(r3.iloc[col_idx["disc1"]]),  to_int(r3.iloc[col_idx["disc2"]])

        details.append(f"\n### 🏫 {sname}")

        # --- 점검 1: 개학일 비교 ---
        details.append("\n**[점검1] 개학일 비교**")
        if open_12 and open_3:
            if open_12 == open_3: details.append(f"- ✅ 개학일 동일 ({open_12})")
            else:
                msg = f"❌ 개학일 불일치: 1~2학년({open_12}) vs 3학년({open_3})"
                details.append(f"- {msg}"); errors.append(f"[{sname}] {msg}")
        else:
            msg = f"⚠️ 개학일 데이터 누락"
            details.append(f"- {msg}"); errors.append(f"[{sname}] {msg}")

        # --- 점검 2: 1학기 수업일수 ---
        for label, od, sc, s1d, d1 in [("1~2학년", open_12, s_close_12, sem1_12, disc1_12), ("3학년", open_3, s_close_3, sem1_3, disc1_3)]:
            details.append(f"\n**[점검2] {label} 1학기 수업일수**")
            if od and sc:
                wd = count_weekdays(od, sc)
                hol = count_holidays_in_range(od, sc, holidays_dict)
                calc = wd - hol - d1
                details.append(f"- 평일 {wd}일 − 공휴일 {hol}일 − 재량휴업 {d1}일 = **{calc}일** (기재: {s1d}일)")
                if calc == s1d: details.append("- ✅ 일치")
                else:
                    msg = f"❌ {label} 1학기 불일치: 계산 {calc}일 vs 기재 {s1d}일"
                    details.append(f"- {msg}"); errors.append(f"[{sname}] {msg}")
            else:
                details.append(f"- ⚠️ 1학기 날짜 누락")

        # --- 점검 3: 2학기 수업일수 ---
        for label, so, wc, wo, ed, s2d, d2 in [("1~2학년", s_open_12, w_close_12, w_open_12, end_12, sem2_12, disc2_12), ("3학년", s_open_3, w_close_3, w_open_3, grad_3, sem2_3, disc2_3)]:
            details.append(f"\n**[점검3] {label} 2학기 수업일수**")
            if so and wc:
                wd_main = count_weekdays(so, wc)
                hol_main = count_holidays_in_range(so, wc, holidays_dict)
                calc = wd_main - hol_main - d2

                if wo is not None and ed is not None:
                    wd_ex = count_weekdays(wo, ed)
                    hol_ex = count_holidays_in_range(wo, ed, holidays_dict)
                    calc += (wd_ex - hol_ex)

                details.append(f"- 총 계산: **{calc}일** (기재: {s2d}일)")
                if calc == s2d: details.append("- ✅ 일치")
                else:
                    msg = f"❌ {label} 2학기 불일치: 계산 {calc}일 vs 기재 {s2d}일"
                    details.append(f"- {msg}"); errors.append(f"[{sname}] {msg}")
            else:
                details.append(f"- ⚠️ 2학기 날짜 누락")

        # --- 점검 4: 재량휴업일 적합성 점검 ---
        details.append("\n**[점검4] 재량휴업일 적합성 점검 (공휴일/주말 중복 여부)**")
        
        def get_valid_dates(row, col_list

# ============================================================
# 6. Streamlit UI (윤년 UI 연동 반영)
# ============================================================
st.title("📋 중, 고등학교 학사일정 점검 시스템")
st.markdown("---")

with st.sidebar:
    st.header("⚙️ 설정")

    current_year = date.today().year
    school_year = st.number_input(
        "📅 학년도 선택", min_value=2020, max_value=2040,
        value=current_year,
        help="학사일정 범위: 선택연도 3월 1일 ~ 다음 연도 2월 말일"
    )
    
    # [보완] 학년도 종료일 자동 계산 (윤년 반영)
    school_year_start = date(school_year, 3, 1)
    school_year_end = date(school_year + 1, 3, 1) - timedelta(days=1)
    st.caption(f"📆 적용 범위: **{school_year_start.strftime('%Y.%m.%d')} ~ {school_year_end.strftime('%Y.%m.%d')}**")

    st.markdown("---")
    st.header("📅 공휴일 관리")

    if "holidays" not in st.session_state or st.session_state.get("last_year") != school_year:
        try:
            st.session_state.holidays = generate_holidays(school_year)
            st.session_state.last_year = school_year
        except Exception as e:
            st.error(f"공휴일 자동 생성 실패: {e}")
            st.session_state.holidays = {}
            st.session_state.last_year = school_year

    if st.button("🔄 공휴일 초기화 (자동 재생성)", use_container_width=True):
        try:
            st.session_state.holidays = generate_holidays(school_year)
            st.session_state.last_year = school_year
            st.rerun()
        except Exception as e:
            st.error(f"생성 실패: {e}")

    st.subheader(f"현재 공휴일 ({len(st.session_state.holidays)}일)")
    dow_names = ["월","화","수","목","금","토","일"]

    to_delete = None
    for d in sorted(st.session_state.holidays.keys()):
        n = st.session_state.holidays[d]
        dow = dow_names[d.weekday()]
        col1, col2 = st.columns([4, 1])
        col1.write(f"{d.strftime('%m/%d')}({dow}) {n}")
        if col2.button("🗑", key=f"del_{d}", help=f"{n} 삭제"):
            to_delete = d

    if to_delete:
        del st.session_state.holidays[to_delete]
        st.rerun()

    st.markdown("---")
    st.subheader("➕ 공휴일 수동 추가")
    st.caption("선거일 등 임시 공휴일을 추가하세요")

    add_date = st.date_input(
        "날짜",
        value=date(school_year, 6, 1),
        min_value=school_year_start,
        max_value=school_year_end,  # [보완] 동적 종료일 반영
        key="add_hol_date"
    )
    add_name = st.text_input("공휴일 이름", value="", placeholder="예: 지방선거일", key="add_hol_name")

    if st.button("✅ 추가", use_container_width=True, key="add_hol_btn"):
        if add_name.strip():
            st.session_state.holidays[add_date] = add_name.strip()
            st.success(f"{add_date} {add_name} 추가됨")
            st.rerun()
        else:
            st.warning("공휴일 이름을 입력하세요")

    st.markdown("---")
    st.header("📖 점검 기준")
    st.markdown("""
    **1. 개학일 비교**
    - 1~2학년과 3학년의 개학일 동일 여부

    **2. 1학기 수업일수**
    - (여름방학식 − 개학일 평일) − 공휴일 − 재량휴업일

    **3. 2학기 수업일수**
    - (겨울방학식 − 여름개학식 평일) − 공휴일 − 재량휴업일
    - 겨울개학식 있으면: + (종업/졸업식 − 겨울개학식)
    """)

# ------ 메인 영역 ------
holidays_dict = st.session_state.holidays

st.header("📂 파일 업로드")
st.info(f"💡 **{school_year}학년도** 학사일정 파일을 업로드하세요. 여러 학교 파일을 동시에 올릴 수 있습니다.")

uploaded_files = st.file_uploader(
    "학사일정 엑셀 파일을 선택하세요",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    st.markdown("---")
    st.header("🔍 점검 결과")

    total_errors = 0
    all_results = {}

    for uf in uploaded_files:
        try:
            df = pd.read_excel(uf, sheet_name=0, header=None)
            # [변경] school_year 파라미터를 함께 전달
            errs, details = check_school(uf.name, df, holidays_dict, school_year) 
            all_results[uf.name] = {"errors": errs, "details": details}
            total_errors += len(errs)
        except Exception as e:
            all_results[uf.name] = {
                "errors": [f"파일 처리 오류: {str(e)}"],
                "details": [f"\n❌ 파일 읽기 오류: {str(e)}"]
            }
            total_errors += 1

    st.subheader("📊 전체 요약")
    col1, col2, col3 = st.columns(3)
    col1.metric("점검 학교 수", f"{len(all_results)}개")
    col2.metric("총 오류 건수", f"{total_errors}건",
                delta=None if total_errors == 0 else f"{total_errors}건 발견",
                delta_color="off" if total_errors == 0 else "inverse")
    col3.metric("정상 학교 수", f"{sum(1 for v in all_results.values() if not v['errors'])}개")

    if total_errors > 0:
        st.subheader("🔴 오류 요약")
        err_data = []
        for fname, res in all_results.items():
            school = fname.split('_')[0]
            for e in res["errors"]:
                err_data.append({"학교": school, "오류 내용": e})
        if err_data:
            st.dataframe(pd.DataFrame(err_data), use_container_width=True, hide_index=True)
    else:
        st.success("🎉 모든 학교가 점검을 통과했습니다!")

    st.markdown("---")
    st.subheader("📝 학교별 상세 결과")
    for fname, res in all_results.items():
        school = fname.split('_')[0]
        icon = "🟢" if not res["errors"] else "🔴"
        with st.expander(f"{icon} {school} ({len(res['errors'])}건 오류)", expanded=bool(res["errors"])):
            for line in res["details"]:
                st.markdown(line)
else:
    st.warning("👆 위에서 학사일정 엑셀 파일을 업로드해 주세요.")
    with st.expander("📖 사용 방법 안내", expanded=True):
        st.markdown("""
        ### 사용 순서
        1. 좌측 사이드바에서 **학년도** 선택 → 공휴일 자동 생성
        2. 필요시 **공휴일 추가/삭제** (선거일 등)
        3. 엑셀 파일 **업로드** → 자동 점검

        ### 파일 형식
        - **파일명**: `학교명_XXXX학년도 고등학교 학사일정 현황.xlsx`
        - 교육청 배포 양식 그대로 사용

        ### 대체공휴일 규칙
        | 구분 | 규칙 |
        |:---|:---|
        | **설날·추석** | 명절 **당일**이 토/일일 때만 대체공휴일 추가 |
        | **기타 공휴일** | 해당일이 토/일이면 대체공휴일 추가 |
        | **현충일** | 대체공휴일 미적용 |
        """)
