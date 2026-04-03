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
    """
    일반 공휴일(1일짜리)의 대체공휴일 계산
    - 해당 날이 토/일이면 다음 첫 평일을 대체공휴일로 추가
    """
    result = [(holiday_date, name)]
    if no_substitute:
        return result

    if holiday_date.weekday() >= 5:  # 토 or 일
        occupied = all_dates | {holiday_date}
        sub = next_weekday_not_in(holiday_date + timedelta(days=1), occupied)
        result.append((sub, f"{name} 대체공휴일"))

    return result

def apply_substitute_holiday3(center_date, name, all_dates):
    """
    명절(3일 연휴: 설날/추석)의 대체공휴일 계산
    ★ 핵심 규칙: 명절 '당일'이 토요일 또는 일요일일 때만 대체공휴일 1일 추가
    - 연휴 3일(전날/당일/다음날)은 무조건 공휴일로 등록
    """
    day_before = center_date - timedelta(days=1)
    day_after  = center_date + timedelta(days=1)
    base_days  = [day_before, center_date, day_after]

    result = [(d, f"{name} 연휴") for d in base_days]

    # 당일이 토요일 또는 일요일인 경우에만 대체공휴일 추가
    if center_date.weekday() >= 5:
        occupied = all_dates | set(base_days)
        sub = next_weekday_not_in(day_after + timedelta(days=1), occupied)
        result.append((sub, f"{name} 대체공휴일"))

    return result

# ============================================================
# 3. 학년도 기준 공휴일 자동 생성
# ============================================================
def generate_holidays(school_year):
    """
    학년도(school_year) 기준 공휴일 자동 생성
    범위: school_year.3.1 ~ (school_year+1).2.28
    """
    Y = school_year
    all_dates = set()
    holidays = []

    # --- 음력 공휴일 ---
    buddha         = lunar_to_solar(Y, 4, 8)
    chuseok_center = lunar_to_solar(Y, 8, 15)
    seollal_center = lunar_to_solar(Y + 1, 1, 1)

    # --- 양력 고정 공휴일 (일반: 당일이 토/일이면 대체공휴일) ---
    fixed = [
        (date(Y, 3, 1),     "삼일절",     False),
        (date(Y, 5, 5),     "어린이날",   False),
        (date(Y, 6, 6),     "현충일",     True),   # 대체공휴일 미적용
        (date(Y, 8, 15),    "광복절",     False),
        (date(Y, 10, 3),    "개천절",     False),
        (date(Y, 10, 9),    "한글날",     False),
        (date(Y, 12, 25),   "기독탄신일", False),
        (date(Y + 1, 1, 1), "신정",       False),
    ]
    if Y >= 2026:
        fixed.append((date(Y, 7, 17), "제헌절", False))

    # ① 일반 공휴일 처리 (1일짜리)
    for hdate, hname, no_sub in fixed:
        entries = apply_substitute_single(hdate, hname, all_dates, no_sub)
        for d, n in entries:
            all_dates.add(d)
            holidays.append((d, n))

    # ② 부처님오신날 (1일짜리, 대체공휴일 적용)
    entries = apply_substitute_single(buddha, "부처님오신날", all_dates, False)
    for d, n in entries:
        all_dates.add(d)
        holidays.append((d, n))

    # ③ 추석 (3일 연휴, 당일 기준 대체공휴일)
    entries = apply_substitute_holiday3(chuseok_center, "추석", all_dates)
    for d, n in entries:
        all_dates.add(d)
        holidays.append((d, n))

    # ④ 설날 (3일 연휴, 당일 기준 대체공휴일)
    entries = apply_substitute_holiday3(seollal_center, "설날", all_dates)
    for d, n in entries:
        all_dates.add(d)
        holidays.append((d, n))

    # 학년도 범위 필터: Y.3.1 ~ (Y+1).2.28
    start = date(Y, 3, 1)
    end   = date(Y + 1, 2, 28)
    holidays = [(d, n) for d, n in holidays if start <= d <= end]

    # 평일만 필터 (토/일은 수업일 계산에서 이미 제외됨)
    holidays = [(d, n) for d, n in holidays if d.weekday() < 5]

    return dict(sorted(holidays, key=lambda x: x[0]))

# ============================================================
# 4. 유틸리티 함수
# ============================================================
def count_weekdays(start_date, end_date):
    if not start_date or not end_date or start_date > end_date:
        return 0
    count = 0
    current = start_date
    while current <= end_date:
        if current.weekday() < 5:
            count += 1
        current += timedelta(days=1)
    return count

def count_holidays_in_range(start_date, end_date, holidays_set):
    return sum(1 for h in holidays_set if start_date <= h <= end_date and h.weekday() < 5)

def to_date(val):
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, pd.Timestamp):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        val = val.strip()
        if val == '' or val.lower() == 'nat':
            return None
        try:
            return pd.to_datetime(val).date()
        except:
            return None
    return None

def to_int(val):
    if val is None:
        return 0
    try:
        if pd.isna(val):
            return 0
    except (TypeError, ValueError):
        pass
    try:
        return int(float(val))
    except:
        return 0

# ============================================================
# 5. 핵심 점검 함수
# ============================================================
def check_school(file_name, df, holidays_set):
    school_name = os.path.basename(file_name).split('_')[0]
    errors, details = [], []

    rows_12, rows_3 = [], []
    for idx in range(df.shape[0]):
        row = df.iloc[idx]
        grade = str(row.iloc[12]).strip() if pd.notna(row.iloc[12]) else ''
        serial = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        if serial == '예시':
            continue
        if grade == '1~2':
            rows_12.append(row)
        elif grade == '3':
            rows_3.append(row)

    if not rows_12 or not rows_3:
        errors.append(f"[{school_name}] 데이터 행(1~2학년/3학년)을 찾을 수 없습니다.")
        return errors, details

    for r12, r3 in zip(rows_12, rows_3):
        sname = str(r12.iloc[6]) if pd.notna(r12.iloc[6]) else school_name

        open_12      = to_date(r12.iloc[16])
        open_3       = to_date(r3.iloc[16])
        s_close_12   = to_date(r12.iloc[17])
        s_open_12    = to_date(r12.iloc[18])
        w_close_12   = to_date(r12.iloc[19])
        w_open_12    = to_date(r12.iloc[20])
        end_12       = to_date(r12.iloc[21])
        s_close_3    = to_date(r3.iloc[17])
        s_open_3     = to_date(r3.iloc[18])
        w_close_3    = to_date(r3.iloc[19])
        w_open_3     = to_date(r3.iloc[20])
        grad_3       = to_date(r3.iloc[22])
        if grad_3 is None:
            grad_3   = to_date(r3.iloc[21])

        sem1_12, sem2_12   = to_int(r12.iloc[13]), to_int(r12.iloc[14])
        sem1_3,  sem2_3    = to_int(r3.iloc[13]),  to_int(r3.iloc[14])
        disc1_12, disc2_12 = to_int(r12.iloc[51]), to_int(r12.iloc[52])
        disc1_3,  disc2_3  = to_int(r3.iloc[51]),  to_int(r3.iloc[52])

        details.append(f"\n### 🏫 {sname}")

        # --- 점검 1: 개학일 비교 ---
        details.append("\n**[점검1] 개학일 비교**")
        if open_12 and open_3:
            if open_12 == open_3:
                details.append(f"- ✅ 개학일 동일 ({open_12})")
            else:
                msg = f"❌ 개학일 불일치: 1~2학년({open_12}) vs 3학년({open_3})"
                details.append(f"- {msg}")
                errors.append(f"[{sname}] {msg}")
        else:
            msg = f"⚠️ 개학일 데이터 누락 (1~2학년={open_12}, 3학년={open_3})"
            details.append(f"- {msg}")
            errors.append(f"[{sname}] {msg}")

        # --- 점검 2: 1학기 수업일수 ---
        for label, od, sc, s1d, d1 in [
            ("1~2학년", open_12, s_close_12, sem1_12, disc1_12),
            ("3학년",   open_3,  s_close_3,  sem1_3,  disc1_3),
        ]:
            details.append(f"\n**[점검2] {label} 1학기 수업일수**")
            if od and sc:
                wd = count_weekdays(od, sc)
                hol = count_holidays_in_range(od, sc, holidays_set)
                calc = wd - hol - d1
                details.append(f"- 개학일({od}) → 여름방학식({sc})")
                details.append(f"- 평일 {wd}일 − 공휴일 {hol}일 − 재량휴업 {d1}일 = **{calc}일**")
                details.append(f"- 기재 수업일수: **{s1d}일**")
                if calc == s1d:
                    details.append("- ✅ 일치")
                else:
                    msg = f"❌ {label} 1학기 불일치: 계산 {calc}일 vs 기재 {s1d}일 (차이 {calc - s1d:+d}일)"
                    details.append(f"- {msg}")
                    errors.append(f"[{sname}] {msg}")
            else:
                msg = f"⚠️ {label} 1학기 날짜 누락"
                details.append(f"- {msg}")
                errors.append(f"[{sname}] {msg}")

        # --- 점검 3: 2학기 수업일수 ---
        for label, so, wc, wo, ed, s2d, d2 in [
            ("1~2학년", s_open_12, w_close_12, w_open_12, end_12,  sem2_12, disc2_12),
            ("3학년",   s_open_3,  w_close_3,  w_open_3,  grad_3,  sem2_3,  disc2_3),
        ]:
            details.append(f"\n**[점검3] {label} 2학기 수업일수**")
            if so and wc:
                wd_main = count_weekdays(so, wc)
                hol_main = count_holidays_in_range(so, wc, holidays_set)
                calc = wd_main - hol_main - d2
                details.append(f"- 여름개학식({so}) → 겨울방학식({wc})")
                details.append(f"- 기본구간: 평일 {wd_main}일 − 공휴일 {hol_main}일 − 재량휴업 {d2}일 = {calc}일")

                if wo is not None and ed is not None:
                    wd_ex = count_weekdays(wo, ed)
                    hol_ex = count_holidays_in_range(wo, ed, holidays_set)
                    extra = wd_ex - hol_ex
                    calc += extra
                    end_label = '종업식' if label == '1~2학년' else '졸업식'
                    details.append(f"- 겨울개학식({wo}) → {end_label}({ed}): +{extra}일")
                    details.append(f"- 총 계산: **{calc}일**")
                else:
                    details.append("- 겨울개학식: 공란 → 추가 계산 없음")
                    details.append(f"- 총 계산: **{calc}일**")

                details.append(f"- 기재 수업일수: **{s2d}일**")
                if calc == s2d:
                    details.append("- ✅ 일치")
                else:
                    msg = f"❌ {label} 2학기 불일치: 계산 {calc}일 vs 기재 {s2d}일 (차이 {calc - s2d:+d}일)"
                    details.append(f"- {msg}")
                    errors.append(f"[{sname}] {msg}")
            else:
                msg = f"⚠️ {label} 2학기 날짜 누락"
                details.append(f"- {msg}")
                errors.append(f"[{sname}] {msg}")

    return errors, details

# ============================================================
# 6. Streamlit UI
# ============================================================
st.title("📋 중, 고등학교 학사일정 점검 시스템")
st.markdown("---")

with st.sidebar:
    st.header("⚙️ 설정")

    current_year = date.today().year
    school_year = st.number_input(
        "📅 학년도 선택", min_value=2020, max_value=2040,
        value=current_year,
        help=f"학사일정 범위: 선택연도.3.1 ~ 다음연도.2.28"
    )
    st.caption(f"📆 적용 범위: **{school_year}.3.1 ~ {school_year+1}.2.28**")

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
        min_value=date(school_year, 3, 1),
        max_value=date(school_year + 1, 2, 28),
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
holidays_set = set(st.session_state.holidays.keys())

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
            errs, details = check_school(uf.name, df, holidays_set)
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
    col3.metric("정상 학교 수",
                f"{sum(1 for v in all_results.values() if not v['errors'])}개")

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
        st.markdown(f"""
        ### 사용 순서
        1. 좌측 사이드바에서 **학년도** 선택 → 공휴일 자동 생성
        2. 필요시 **공휴일 추가/삭제** (선거일 등)
        3. 엑셀 파일 **업로드** → 자동 점검

        ### 파일 형식
        - **파일명**: `학교명_XXXX학년도 고등학교 학사일정 현황.xlsx`
        - 교육청 배포 양식 그대로 사용

        ### 점검 항목
        | # | 점검 내용 | 설명 |
        |:-:|:---|:---|
        | 1 | 개학일 비교 | 1~2학년과 3학년 동일 여부 |
        | 2 | 1학기 수업일수 | 개학일~여름방학식 평일 − 공휴일 − 재량휴업일 |
        | 3 | 2학기 수업일수 | 여름개학식~겨울방학식 평일 − 공휴일 − 재량휴업일 + 겨울개학 후 구간 |

        ### 대체공휴일 규칙
        | 구분 | 규칙 |
        |:---|:---|
        | **설날·추석** | 명절 **당일**이 토/일일 때만 대체공휴일 추가 |
        | **기타 공휴일** | 해당일이 토/일이면 대체공휴일 추가 |
        | **현충일** | 대체공휴일 미적용 |
        """)
