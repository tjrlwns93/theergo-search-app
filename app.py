# -*- coding: utf-8 -*-
"""
더에르고 검색 웹앱 (Streamlit)
- 비밀번호 로그인 후에만 화면이 보인다.
- [검색어]를 넣으면 theergo.co.kr 첫 번째 결과의 제목/링크를 가져온다.
- 결과를 Supabase(PostgreSQL)에 저장하고, 저장 이력을 보여준다.
- 접속정보/비밀번호는 코드가 아니라 secrets(.streamlit/secrets.toml)에서 읽는다.
"""
import datetime
import html as html_mod
import io
import json
import re
import time
import urllib.parse
import urllib.request

import openpyxl
import psycopg2
import streamlit as st

# ──────────────────────────────────────────────────────────────
# 설정 (secrets 에서 읽음 — 코드에 비밀정보를 두지 않는다)
# ──────────────────────────────────────────────────────────────
BASE = "https://theergo.co.kr"
SEARCH_URL = BASE + "/product/search.html?keyword="
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Search-Bot"}

# 연습용 주문 API (인증 불필요, date/page 파라미터, 100건씩 페이지네이션)
ORDERS_API = "https://sp.ermore.co.kr/api/edu-leader/practice/orders"

# 날씨 API (Open-Meteo, 무료·키 불필요) — 서울 기준
WEATHER_API = "https://api.open-meteo.com/v1/forecast"
SEOUL_LAT, SEOUL_LON = 37.5665, 126.9780

st.set_page_config(page_title="더에르고 검색기", page_icon="🔎")


def get_secret(key: str, default=None):
    try:
        return st.secrets[key]
    except Exception:
        return default


# ──────────────────────────────────────────────────────────────
# 1) 비밀번호 로그인 게이트
# ──────────────────────────────────────────────────────────────
def check_password() -> bool:
    if st.session_state.get("auth_ok"):
        return True

    st.title("🔒 로그인")
    st.caption("비밀번호를 입력해야 화면이 보입니다.")
    pw = st.text_input("비밀번호", type="password")
    if st.button("로그인"):
        real = get_secret("app_password")
        if real and pw == real:
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("비밀번호가 틀렸습니다.")
    return False


# ──────────────────────────────────────────────────────────────
# 2) 데이터베이스 (Supabase PostgreSQL)
# ──────────────────────────────────────────────────────────────
def get_conn():
    # 방법 1 (권장): 항목별 접속정보 — 비번 특수문자/URL 인코딩 걱정 없음
    host = get_secret("db_host")
    if host:
        return psycopg2.connect(
            host=host,
            port=int(get_secret("db_port", 5432)),
            user=get_secret("db_user"),
            password=get_secret("db_password"),
            dbname=get_secret("db_name", "postgres"),
            sslmode="require",
            connect_timeout=15,
        )

    # 방법 2: 전체 연결 URL (특수문자는 직접 URL 인코딩 필요)
    url = get_secret("db_url")
    if not url:
        raise RuntimeError(
            "secrets 에 db_host(항목별) 또는 db_url 이 필요합니다."
        )
    if "<region>" in url or "[YOUR-PASSWORD]" in url or "[" in url:
        raise RuntimeError(
            "db_url 에 <region> / [YOUR-PASSWORD] 같은 자리표시자가 그대로 있습니다. "
            "실제 값으로 바꿔주세요."
        )
    if "sslmode=" not in url:
        url += ("&" if "?" in url else "?") + "sslmode=require"
    return psycopg2.connect(url, connect_timeout=15)


def init_db():
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS search_results (
                id BIGSERIAL PRIMARY KEY,
                keyword TEXT NOT NULL,
                title TEXT,
                link TEXT,
                searched_at TIMESTAMPTZ NOT NULL DEFAULT now()
            );
            """
        )
        # 주문 테이블: order_no 를 PK 로 두어 같은 날짜 재수집 시 중복 방지
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS orders (
                order_no       TEXT PRIMARY KEY,
                order_date     DATE NOT NULL,
                ordered_at     TEXT,
                product        TEXT,
                product_option TEXT,
                qty            INTEGER,
                amount         BIGINT,
                status         TEXT,
                collected_at   TIMESTAMPTZ NOT NULL DEFAULT now()
            );
            """
        )
        cur.execute(
            "CREATE INDEX IF NOT EXISTS idx_orders_date ON orders (order_date);"
        )
        # 날씨 테이블: (도시, 날짜) 를 PK 로 두어 같은 날짜 재수집 시 중복 방지
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS weather (
                city         TEXT NOT NULL DEFAULT 'Seoul',
                weather_date DATE NOT NULL,
                temp_max     REAL,
                temp_min     REAL,
                precipitation REAL,
                collected_at TIMESTAMPTZ NOT NULL DEFAULT now(),
                PRIMARY KEY (city, weather_date)
            );
            """
        )
        conn.commit()


def save_results(rows):
    """rows: list of (keyword, title, link)"""
    with get_conn() as conn, conn.cursor() as cur:
        cur.executemany(
            "INSERT INTO search_results (keyword, title, link) VALUES (%s, %s, %s)",
            rows,
        )
        conn.commit()


def fetch_history(limit=200):
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            "SELECT keyword, title, link, searched_at "
            "FROM search_results ORDER BY id DESC LIMIT %s",
            (limit,),
        )
        return cur.fetchall()


# ──────────────────────────────────────────────────────────────
# 2-1) 주문 수집 (연습용 API → DB)
# ──────────────────────────────────────────────────────────────
def fetch_orders(date_str, progress=None, max_pages=1000):
    """해당 날짜의 모든 주문을 page 를 넘기며(100건씩) 모아서 반환.
    반환: (rows_list, total). 종료조건: 빈 페이지 또는 total 도달."""
    all_rows = []
    total = None
    page = 1
    while page <= max_pages:
        url = ORDERS_API + "?" + urllib.parse.urlencode({"date": date_str, "page": page})
        req = urllib.request.Request(
            url, headers={"User-Agent": HEADERS["User-Agent"], "Accept": "application/json"}
        )
        with urllib.request.urlopen(req, timeout=25) as resp:
            data = json.loads(resp.read().decode("utf-8", "replace"))
        rows = data.get("rows") or []
        if not rows:
            break
        all_rows.extend(rows)
        total = data.get("total")
        if progress:
            progress(len(all_rows), total, page)
        if total is not None and len(all_rows) >= total:
            break
        page += 1
    return all_rows, total


def upsert_orders(date_obj, rows):
    """order_no 기준 UPSERT. 같은 주문을 다시 넣어도 중복으로 쌓이지 않고 갱신됨.
    반환: (신규 건수, 갱신 건수)."""
    if not rows:
        return 0, 0
    new_cnt = upd_cnt = 0
    with get_conn() as conn, conn.cursor() as cur:
        for r in rows:
            cur.execute(
                """
                INSERT INTO orders
                    (order_no, order_date, ordered_at, product, product_option,
                     qty, amount, status, collected_at)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, now())
                ON CONFLICT (order_no) DO UPDATE SET
                    order_date     = EXCLUDED.order_date,
                    ordered_at     = EXCLUDED.ordered_at,
                    product        = EXCLUDED.product,
                    product_option = EXCLUDED.product_option,
                    qty            = EXCLUDED.qty,
                    amount         = EXCLUDED.amount,
                    status         = EXCLUDED.status,
                    collected_at   = now()
                RETURNING (xmax = 0) AS inserted
                """,
                (
                    r.get("order_no"), date_obj, r.get("ordered_at"),
                    r.get("product"), r.get("option"), r.get("qty"),
                    r.get("amount"), r.get("status"),
                ),
            )
            if cur.fetchone()[0]:
                new_cnt += 1
            else:
                upd_cnt += 1
        conn.commit()
    return new_cnt, upd_cnt


def fetch_orders_by_date(date_obj, limit=2000):
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            "SELECT order_no, ordered_at, product, product_option, qty, amount, status "
            "FROM orders WHERE order_date = %s ORDER BY order_no LIMIT %s",
            (date_obj, limit),
        )
        return cur.fetchall()


# ──────────────────────────────────────────────────────────────
# 2-2) 날씨 수집 (Open-Meteo → DB), 서울 기준
# ──────────────────────────────────────────────────────────────
def fetch_weather(start_date, end_date):
    """서울의 날짜별 최고/최저기온·강수량을 가져와 리스트로 반환.
    반환: [(date, tmax, tmin, precip), ...]"""
    url = WEATHER_API + "?" + urllib.parse.urlencode({
        "latitude": SEOUL_LAT,
        "longitude": SEOUL_LON,
        "daily": "temperature_2m_max,temperature_2m_min,precipitation_sum",
        "timezone": "Asia/Seoul",
        "start_date": start_date,
        "end_date": end_date,
    })
    req = urllib.request.Request(
        url, headers={"User-Agent": HEADERS["User-Agent"], "Accept": "application/json"}
    )
    with urllib.request.urlopen(req, timeout=25) as resp:
        data = json.loads(resp.read().decode("utf-8", "replace"))
    daily = data.get("daily") or {}
    times = daily.get("time") or []
    tmax = daily.get("temperature_2m_max") or []
    tmin = daily.get("temperature_2m_min") or []
    prcp = daily.get("precipitation_sum") or []
    rows = []
    for i, d in enumerate(times):
        rows.append((
            d,
            tmax[i] if i < len(tmax) else None,
            tmin[i] if i < len(tmin) else None,
            prcp[i] if i < len(prcp) else None,
        ))
    return rows


def upsert_weather(rows, city="Seoul"):
    """(도시,날짜) 기준 UPSERT. 같은 날짜를 다시 가져와도 중복 없이 갱신.
    반환: (신규 건수, 갱신 건수)."""
    if not rows:
        return 0, 0
    new_cnt = upd_cnt = 0
    with get_conn() as conn, conn.cursor() as cur:
        for d, tmax, tmin, prcp in rows:
            cur.execute(
                """
                INSERT INTO weather
                    (city, weather_date, temp_max, temp_min, precipitation, collected_at)
                VALUES (%s, %s, %s, %s, %s, now())
                ON CONFLICT (city, weather_date) DO UPDATE SET
                    temp_max      = EXCLUDED.temp_max,
                    temp_min      = EXCLUDED.temp_min,
                    precipitation = EXCLUDED.precipitation,
                    collected_at  = now()
                RETURNING (xmax = 0) AS inserted
                """,
                (city, d, tmax, tmin, prcp),
            )
            if cur.fetchone()[0]:
                new_cnt += 1
            else:
                upd_cnt += 1
        conn.commit()
    return new_cnt, upd_cnt


def fetch_weather_saved(city="Seoul", limit=365):
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            "SELECT weather_date, temp_max, temp_min, precipitation "
            "FROM weather WHERE city = %s ORDER BY weather_date DESC LIMIT %s",
            (city, limit),
        )
        return cur.fetchall()


# ──────────────────────────────────────────────────────────────
# 3) 검색 (theergo.co.kr)
# ──────────────────────────────────────────────────────────────
def fetch(keyword: str) -> str:
    url = SEARCH_URL + urllib.parse.quote(keyword)
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=25) as resp:
        return resp.read().decode("utf-8", "replace")


def strip_tags(fragment: str) -> str:
    text = re.sub(r"<[^>]+>", " ", fragment)
    text = html_mod.unescape(text)
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"^상품명\s*:\s*", "", text)
    return text


def parse_first_result(page_html: str):
    start = page_html.find("prdList")
    if start == -1:
        return None, None
    region = page_html[start:]
    m = re.search(
        r'<strong class="name">\s*<a\s+href="([^"]+)"[^>]*>(.*?)</a>',
        region,
        re.S,
    )
    if not m:
        return None, None
    href = html_mod.unescape(m.group(1).strip())
    title = strip_tags(m.group(2))
    link = urllib.parse.urljoin(BASE, href)
    return (title or None), link


def search_one(keyword: str):
    try:
        title, link = parse_first_result(fetch(keyword))
        if title is None:
            return "검색결과 없음", ""
        return title, link
    except Exception as e:
        return f"오류: {e}", ""


def to_excel_bytes(rows, header):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "검색결과"
    ws.append(header)
    for r in rows:
        ws.append(list(r))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────
# 4) 메인 화면
# ──────────────────────────────────────────────────────────────
def render_search():
    st.write("검색어를 넣으면 theergo.co.kr 첫 번째 결과의 제목·링크를 가져와 저장합니다.")

    st.subheader("1. 검색어 입력")
    tab1, tab2 = st.tabs(["직접 입력", "엑셀 업로드"])

    keywords = []
    with tab1:
        text = st.text_area(
            "한 줄에 하나씩 입력",
            value="잠옷\n매트\n치약\n스티커",
            height=150,
        )
        if text.strip():
            keywords = [ln.strip() for ln in text.splitlines() if ln.strip()]

    with tab2:
        up = st.file_uploader("[검색어] 열이 있는 .xlsx", type=["xlsx"])
        col = st.text_input("검색어 열 이름", value="검색어")
        if up is not None:
            wb = openpyxl.load_workbook(up)
            ws = wb.active
            header = [c.value for c in ws[1]]
            if col in header:
                idx = header.index(col)
                keywords = [
                    str(row[idx]).strip()
                    for row in ws.iter_rows(min_row=2, values_only=True)
                    if idx < len(row) and row[idx] and str(row[idx]).strip()
                ]
                st.success(f"{len(keywords)}개 검색어를 읽었습니다: {keywords}")
            else:
                st.error(f"'{col}' 열이 없습니다. 헤더: {header}")

    st.subheader("2. 검색 실행")
    if st.button("🚀 검색하고 저장하기", type="primary", disabled=not keywords):
        rows = []
        prog = st.progress(0.0)
        area = st.empty()
        for i, kw in enumerate(keywords, 1):
            title, link = search_one(kw)
            rows.append((kw, title, link))
            area.write(f"[{i}/{len(keywords)}] **{kw}** → {title}")
            prog.progress(i / len(keywords))
            if i < len(keywords):
                time.sleep(1.0)
        try:
            save_results(rows)
            st.success(f"완료! {len(rows)}건을 Supabase에 저장했습니다.")
        except Exception as e:
            st.error(f"저장 실패: {e}")
        st.dataframe(
            [{"검색어": r[0], "제목": r[1], "링크": r[2]} for r in rows],
            use_container_width=True,
        )
        st.download_button(
            "⬇️ 엑셀로 내려받기",
            data=to_excel_bytes(rows, ["검색어", "제목", "링크"]),
            file_name="검색결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.subheader("3. 저장 이력 (최근 200건)")
    if st.button("🔄 이력 새로고침"):
        st.session_state["_reload"] = True
    try:
        hist = fetch_history()
        if hist:
            st.dataframe(
                [
                    {
                        "검색어": h[0],
                        "제목": h[1],
                        "링크": h[2],
                        "시각": h[3].strftime("%Y-%m-%d %H:%M:%S"),
                    }
                    for h in hist
                ],
                use_container_width=True,
            )
        else:
            st.info("아직 저장된 이력이 없습니다.")
    except Exception as e:
        st.error(f"이력 조회 실패: {e}")


def render_orders():
    st.write("날짜를 고르고 버튼을 누르면 그 날 주문을 API에서 100건씩 모두 가져와 DB에 쌓습니다.")

    picked = st.date_input("수집할 날짜", value=datetime.date(2026, 8, 18))
    date_str = picked.strftime("%Y-%m-%d")

    if st.button("📦 이 날짜 주문 수집", type="primary"):
        prog = st.progress(0.0)
        area = st.empty()

        def on_prog(got, total, page):
            if total:
                prog.progress(min(got / total, 1.0))
            area.write(f"{page}페이지까지 {got}/{total or '?'}건 가져오는 중…")

        try:
            rows, total = fetch_orders(date_str, progress=on_prog)
        except Exception as e:
            st.error(f"API 호출 실패: {e}")
            st.stop()

        prog.progress(1.0)
        if not rows:
            st.info(f"{date_str} 에는 주문이 없습니다.")
        else:
            try:
                new_cnt, upd_cnt = upsert_orders(picked, rows)
                st.success(
                    f"완료! {date_str} 주문 {len(rows)}건 수집 "
                    f"→ 신규 {new_cnt}건 · 갱신 {upd_cnt}건 (중복은 자동 제외)"
                )
            except Exception as e:
                st.error(f"저장 실패: {e}")

    st.divider()
    st.caption(f"DB에 저장된 {date_str} 주문")
    try:
        saved = fetch_orders_by_date(picked)
        st.write(f"저장된 주문: **{len(saved)}건**")
        if saved:
            st.dataframe(
                [
                    {
                        "주문번호": s[0],
                        "주문시각": s[1],
                        "상품": s[2],
                        "옵션": s[3],
                        "수량": s[4],
                        "금액": s[5],
                        "상태": s[6],
                    }
                    for s in saved
                ],
                use_container_width=True,
            )
    except Exception as e:
        st.error(f"주문 조회 실패: {e}")


def render_weather():
    st.write("서울의 날짜별 최고·최저 기온과 강수량을 Open-Meteo에서 가져와 DB에 저장합니다.")

    today = datetime.date.today()
    c1, c2 = st.columns(2)
    with c1:
        start = st.date_input("시작 날짜", value=today - datetime.timedelta(days=6))
    with c2:
        end = st.date_input("끝 날짜", value=today)

    if start > end:
        st.warning("시작 날짜가 끝 날짜보다 늦습니다.")
    elif st.button("🌤️ 날씨 가져와 저장", type="primary"):
        try:
            rows = fetch_weather(start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))
        except Exception as e:
            st.error(f"API 호출 실패: {e}")
            st.stop()
        if not rows:
            st.info("해당 기간의 날씨 데이터가 없습니다.")
        else:
            try:
                new_cnt, upd_cnt = upsert_weather(rows)
                st.success(
                    f"완료! {len(rows)}일치 수집 "
                    f"→ 신규 {new_cnt}건 · 갱신 {upd_cnt}건 (중복은 자동 제외)"
                )
            except Exception as e:
                st.error(f"저장 실패: {e}")

    st.divider()
    st.caption("DB에 저장된 서울 날씨 (최근순)")
    try:
        saved = fetch_weather_saved()
        st.write(f"저장된 날짜: **{len(saved)}일**")
        if saved:
            table = [
                {
                    "날짜": s[0].strftime("%Y-%m-%d"),
                    "최고기온(°C)": s[1],
                    "최저기온(°C)": s[2],
                    "강수량(mm)": s[3],
                }
                for s in saved
            ]
            st.dataframe(table, use_container_width=True)
    except Exception as e:
        st.error(f"날씨 조회 실패: {e}")


def main():
    st.title("🔎 더에르고 검색기")

    # DB 준비 (검색결과·주문·날씨 테이블 생성)
    try:
        init_db()
    except Exception as e:
        st.error(f"DB 연결 실패: {e}")
        st.stop()

    tab_search, tab_orders, tab_weather = st.tabs(
        ["🔎 검색어 수집", "📦 주문 수집", "🌤️ 날씨"]
    )
    with tab_search:
        render_search()
    with tab_orders:
        render_orders()
    with tab_weather:
        render_weather()


if not check_password():
    st.stop()

main()
