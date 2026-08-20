# -*- coding: utf-8 -*-
"""
더에르고 검색 웹앱 (Streamlit)
- 비밀번호 로그인 후에만 화면이 보인다.
- [검색어]를 넣으면 theergo.co.kr 첫 번째 결과의 제목/링크를 가져온다.
- 결과를 Supabase(PostgreSQL)에 저장하고, 저장 이력을 보여준다.
- 접속정보/비밀번호는 코드가 아니라 secrets(.streamlit/secrets.toml)에서 읽는다.
"""
import base64
import datetime
import hashlib
import html as html_mod
import io
import json
import os
import re
import secrets as secretsmod
import time
import urllib.parse
import urllib.request

import openpyxl
import psycopg2
import streamlit as st
from dotenv import load_dotenv
from streamlit_oauth import OAuth2Component

load_dotenv()  # 로컬 .env 로드 (Streamlit Cloud 에서는 Secrets 사용)

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
# 1) 인증 · 권한 (구글 로그인 + 아이디/비번 로그인 + 화면 권한)
# ──────────────────────────────────────────────────────────────
def cfg(key, default=None):
    """설정값: .env(os.environ) 우선, 없으면 st.secrets."""
    v = os.getenv(key)
    if v not in (None, ""):
        return v
    for k in (key, key.lower(), key.upper()):
        try:
            val = st.secrets[k]
            if val not in (None, ""):
                return val
        except Exception:
            pass
    return default


def admin_emails():
    raw = cfg("ADMIN_EMAILS",
              "tjrlwns93@ermore.co.kr,rho_js@ermore.co.kr,shheo@ermore.co.kr")
    return [e.strip().lower() for e in str(raw).split(",") if e.strip()]


def allowed_domain():
    return str(cfg("ALLOWED_DOMAIN", "ermore.co.kr")).lower().lstrip("@")


# 비밀번호 해시 (표준 라이브러리 pbkdf2)
def hash_pw(pw):
    salt = secretsmod.token_hex(16)
    dk = hashlib.pbkdf2_hmac("sha256", pw.encode(), salt.encode(), 200_000)
    return f"pbkdf2$200000${salt}${dk.hex()}"


def verify_pw(pw, stored):
    try:
        _algo, iters, salt, h = stored.split("$")
        dk = hashlib.pbkdf2_hmac("sha256", pw.encode(), salt.encode(), int(iters))
        return secretsmod.compare_digest(dk.hex(), h)
    except Exception:
        return False


def _rows_as_dicts(cur):
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, r)) for r in cur.fetchall()]


_USER_COLS = ("id, email, username, name, password_hash, auth_type, "
              "is_admin, is_active, must_change_pw, screens")


def db_get_user_by_username(username):
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(f"SELECT {_USER_COLS} FROM app_users WHERE username=%s", (username,))
        rows = _rows_as_dicts(cur)
    return rows[0] if rows else None


def db_upsert_google_user(email, name):
    is_admin = email.lower() in admin_emails()
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            """
            INSERT INTO app_users (email, name, auth_type, is_admin)
            VALUES (%s, %s, 'google', %s)
            ON CONFLICT (email) DO UPDATE SET name = EXCLUDED.name, is_admin = %s
            RETURNING id, email, username, name, auth_type, is_admin,
                      is_active, must_change_pw, screens
            """,
            (email.lower(), name, is_admin, is_admin),
        )
        conn.commit()
        rows = _rows_as_dicts(cur)
    return rows[0] if rows else None


def _gen_username(name):
    base = re.sub(r"[^a-z0-9]", "", (name or "").lower()) or "user"
    for i in range(0, 1000):
        cand = base if i == 0 else f"{base}{i + 1}"
        if not db_get_user_by_username(cand):
            return cand
    return base + secretsmod.token_hex(3)


def db_create_local_user(name, username=None):
    uname = (username or "").strip() or _gen_username(name)
    temp = secretsmod.token_urlsafe(6)
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            "INSERT INTO app_users (username, name, password_hash, auth_type, must_change_pw) "
            "VALUES (%s, %s, %s, 'local', TRUE)",
            (uname, name, hash_pw(temp)),
        )
        conn.commit()
    return uname, temp


def db_reset_password(user_id):
    temp = secretsmod.token_urlsafe(6)
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            "UPDATE app_users SET password_hash=%s, must_change_pw=TRUE WHERE id=%s",
            (hash_pw(temp), user_id),
        )
        conn.commit()
    return temp


def db_change_password(user_id, new_pw):
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            "UPDATE app_users SET password_hash=%s, must_change_pw=FALSE WHERE id=%s",
            (hash_pw(new_pw), user_id),
        )
        conn.commit()


def db_list_users():
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(f"SELECT {_USER_COLS} FROM app_users ORDER BY id")
        return _rows_as_dicts(cur)


def db_update_user_screens(user_id, screens):
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute("UPDATE app_users SET screens=%s WHERE id=%s", (list(screens), user_id))
        conn.commit()


def db_set_active(user_id, active):
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute("UPDATE app_users SET is_active=%s WHERE id=%s", (bool(active), user_id))
        conn.commit()


def db_list_allowed_emails():
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute("SELECT email FROM allowed_emails ORDER BY email")
        return [r[0] for r in cur.fetchall()]


def db_add_allowed_email(email):
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            "INSERT INTO allowed_emails (email) VALUES (%s) ON CONFLICT DO NOTHING",
            (email.lower(),),
        )
        conn.commit()


def db_remove_allowed_email(email):
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM allowed_emails WHERE lower(email)=lower(%s)", (email,))
        conn.commit()


def db_is_email_allowed(email):
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute("SELECT 1 FROM allowed_emails WHERE lower(email)=lower(%s)", (email,))
        return cur.fetchone() is not None


# ── 구글 로그인 (팝업) ─────────────────────────────────────────
def _google_identity(token):
    """토큰에서 이메일·이름 추출."""
    idt = token.get("id_token")
    if idt:
        try:
            payload = idt.split(".")[1]
            payload += "=" * (-len(payload) % 4)
            data = json.loads(base64.urlsafe_b64decode(payload).decode("utf-8"))
            return (data.get("email") or "").lower(), data.get("name") or ""
        except Exception:
            pass
    at = token.get("access_token")
    if at:
        req = urllib.request.Request(
            "https://www.googleapis.com/oauth2/v3/userinfo",
            headers={"Authorization": f"Bearer {at}"},
        )
        data = json.loads(urllib.request.urlopen(req, timeout=15).read().decode("utf-8"))
        return (data.get("email") or "").lower(), data.get("name") or ""
    return "", ""


def _handle_google_login(email, name):
    email = (email or "").lower()
    if not email:
        st.error("구글 계정 이메일을 확인할 수 없습니다.")
        return
    is_admin = email in admin_emails()
    allowed = (
        is_admin
        or email.endswith("@" + allowed_domain())
        or db_is_email_allowed(email)
    )
    if not allowed:
        st.session_state["denied_email"] = email
        st.rerun()
        return
    row = db_upsert_google_user(email, name)
    screens = SCREEN_KEYS if is_admin else (row.get("screens") or [])
    st.session_state["user"] = {
        "kind": "google", "id": row.get("id"), "email": email,
        "name": name or email, "is_admin": is_admin, "screens": screens,
    }
    st.rerun()


def _handle_local_login(username, password):
    row = db_get_user_by_username((username or "").strip())
    if (not row or row["auth_type"] != "local" or not row["is_active"]
            or not verify_pw(password or "", row["password_hash"] or "")):
        st.error("아이디 또는 비밀번호가 올바르지 않습니다.")
        return
    st.session_state["user"] = {
        "kind": "local", "id": row["id"], "username": row["username"],
        "name": row["name"], "is_admin": False,
        "screens": row["screens"] or [], "must_change_pw": row["must_change_pw"],
    }
    st.rerun()


def render_login():
    if st.session_state.get("denied_email"):
        st.title("🚫 접근 권한이 없습니다")
        st.write(f"**{st.session_state['denied_email']}** 계정은 접근이 허용되지 않았습니다.")
        st.caption(f"@{allowed_domain()} 계정이거나, 관리자가 예외로 허용한 이메일이어야 합니다.")
        if st.button("다른 계정으로 로그인"):
            st.session_state.pop("denied_email", None)
            st.rerun()
        return

    st.title("🔐 로그인")

    gid, gsec = cfg("GOOGLE_CLIENT_ID"), cfg("GOOGLE_CLIENT_SECRET")
    redirect = cfg("APP_URL") or cfg("cafe24_redirect_uri")
    if gid and gsec and redirect:
        oauth2 = OAuth2Component(
            gid, gsec,
            "https://accounts.google.com/o/oauth2/v2/auth",
            "https://oauth2.googleapis.com/token",
            "https://oauth2.googleapis.com/token",
            "https://oauth2.googleapis.com/revoke",
        )
        result = oauth2.authorize_button(
            name="구글 계정으로 로그인",
            redirect_uri=redirect,
            scope="openid email profile",
            key="google_login",
            use_container_width=True,
            extras_params={"prompt": "select_account"},
        )
        if result and "token" in result:
            email, name = _google_identity(result["token"])
            _handle_google_login(email, name)
    else:
        st.info("구글 로그인 설정(GOOGLE_CLIENT_ID / SECRET / APP_URL)이 아직 없습니다.")

    st.divider()
    st.caption("구글이 없는 계정은 아이디·비밀번호로 로그인")
    with st.form("local_login"):
        u = st.text_input("아이디")
        p = st.text_input("비밀번호", type="password")
        if st.form_submit_button("로그인"):
            _handle_local_login(u, p)

    # (임시) 예비 로그인 — 새 로그인 확인 전까지 잠김 방지용. 확인되면 제거 예정.
    bg = cfg("app_password")
    if bg:
        with st.expander("예비 로그인(임시)"):
            bp = st.text_input("예비 비밀번호", type="password", key="bg_pw")
            if st.button("예비 로그인"):
                if bp == bg:
                    st.session_state["user"] = {
                        "kind": "break-glass", "name": "임시관리자",
                        "email": None, "is_admin": True, "screens": SCREEN_KEYS,
                    }
                    st.rerun()
                else:
                    st.error("예비 비밀번호가 틀렸습니다.")


def render_change_pw():
    st.title("🔑 비밀번호 변경")
    st.info("첫 로그인입니다. 새 비밀번호를 설정하세요.")
    with st.form("change_pw"):
        p1 = st.text_input("새 비밀번호(8자 이상)", type="password")
        p2 = st.text_input("새 비밀번호 확인", type="password")
        if st.form_submit_button("변경"):
            if len(p1) < 8:
                st.error("비밀번호는 8자 이상이어야 합니다.")
            elif p1 != p2:
                st.error("두 비밀번호가 다릅니다.")
            else:
                db_change_password(st.session_state["user"]["id"], p1)
                st.session_state["user"]["must_change_pw"] = False
                st.success("변경 완료!")
                st.rerun()


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
        # 카페24 OAuth 토큰 저장 (몰당 1행). 고객정보는 저장하지 않음.
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS cafe24_token (
                mall_id                  TEXT PRIMARY KEY,
                access_token             TEXT,
                refresh_token            TEXT,
                expires_at               TEXT,
                refresh_token_expires_at TEXT,
                updated_at               TIMESTAMPTZ NOT NULL DEFAULT now()
            );
            """
        )
        # 사용자 계정 (구글/로컬) + 화면 권한
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS app_users (
                id             BIGSERIAL PRIMARY KEY,
                email          TEXT UNIQUE,
                username       TEXT UNIQUE,
                name           TEXT,
                password_hash  TEXT,
                auth_type      TEXT NOT NULL DEFAULT 'local',
                is_admin       BOOLEAN NOT NULL DEFAULT FALSE,
                is_active      BOOLEAN NOT NULL DEFAULT TRUE,
                must_change_pw BOOLEAN NOT NULL DEFAULT FALSE,
                screens        TEXT[] NOT NULL DEFAULT '{}',
                created_at     TIMESTAMPTZ NOT NULL DEFAULT now()
            );
            """
        )
        # 도메인 예외로 허용할 이메일 목록
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS allowed_emails (
                email    TEXT PRIMARY KEY,
                added_at TIMESTAMPTZ NOT NULL DEFAULT now()
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


def fetch_weather_collected_today(city="Seoul"):
    """오늘(한국시간) 수집한 날씨만 반환."""
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            "SELECT weather_date, temp_max, temp_min, precipitation "
            "FROM weather WHERE city = %s "
            "AND (collected_at AT TIME ZONE 'Asia/Seoul')::date "
            "    = (now() AT TIME ZONE 'Asia/Seoul')::date "
            "ORDER BY weather_date",
            (city,),
        )
        return cur.fetchall()


# ──────────────────────────────────────────────────────────────
# 2-3) 노션 연동 (Open-Meteo 로 모은 날씨를 노션 DB 에 한 줄씩 추가)
# ──────────────────────────────────────────────────────────────
NOTION_API = "https://api.notion.com/v1"
NOTION_VERSION = "2022-06-28"


def add_weather_rows_to_notion(rows):
    """rows: [(date, tmax, tmin, precip), ...] 를 노션 DB 에 한 줄씩 추가.
    '테스트'(title)=날씨 내용, '담당자'(people)=secrets 의 notion_person_id.
    반환: 추가한 줄 수."""
    token = get_secret("notion_token")
    db_id = get_secret("notion_db_id")
    person_id = get_secret("notion_person_id")
    if not token or not db_id:
        raise RuntimeError("secrets 에 notion_token / notion_db_id 가 필요합니다.")
    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
    }
    added = 0
    for d, tmax, tmin, prcp in rows:
        content = f"{d} 서울 · 최고 {tmax}°C / 최저 {tmin}°C / 강수 {prcp}mm"
        props = {"테스트": {"title": [{"text": {"content": content}}]}}
        if person_id:
            props["담당자"] = {"people": [{"id": person_id}]}
        payload = {"parent": {"database_id": db_id}, "properties": props}
        req = urllib.request.Request(
            NOTION_API + "/pages",
            data=json.dumps(payload).encode("utf-8"),
            headers=headers,
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=25) as resp:
            resp.read()
        added += 1
    return added


# ──────────────────────────────────────────────────────────────
# 2-4) 카페24 관리자 API (OAuth + 토큰 자동갱신). 고객정보 미수집.
# ──────────────────────────────────────────────────────────────
CAFE24_API_VERSION = "2022-06-01"  # X-Cafe24-Api-Version (앱 설정과 맞추세요)


def _cafe24_cfg():
    return (
        get_secret("cafe24_mall_id"),
        get_secret("cafe24_client_id"),
        get_secret("cafe24_client_secret"),
        get_secret("cafe24_redirect_uri"),
    )


def cafe24_authorize_url(state="ergo2-state"):
    """1단계: 승인 요청 주소. scope 는 주문 읽기만(고객정보 scope 없음)."""
    mall_id, client_id, _, redirect_uri = _cafe24_cfg()
    params = {
        "response_type": "code",
        "client_id": client_id,
        "state": state,
        "redirect_uri": redirect_uri,
        "scope": "mall.read_order",
    }
    return (
        f"https://{mall_id}.cafe24api.com/api/v2/oauth/authorize?"
        + urllib.parse.urlencode(params)
    )


def _cafe24_save_token(mall_id, tok):
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            """
            INSERT INTO cafe24_token
                (mall_id, access_token, refresh_token, expires_at,
                 refresh_token_expires_at, updated_at)
            VALUES (%s, %s, %s, %s, %s, now())
            ON CONFLICT (mall_id) DO UPDATE SET
                access_token             = EXCLUDED.access_token,
                refresh_token            = EXCLUDED.refresh_token,
                expires_at               = EXCLUDED.expires_at,
                refresh_token_expires_at = EXCLUDED.refresh_token_expires_at,
                updated_at               = now()
            """,
            (
                mall_id, tok.get("access_token"), tok.get("refresh_token"),
                tok.get("expires_at"), tok.get("refresh_token_expires_at"),
            ),
        )
        conn.commit()


def _cafe24_load_token():
    mall_id = get_secret("cafe24_mall_id")
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(
            "SELECT access_token, refresh_token, expires_at "
            "FROM cafe24_token WHERE mall_id = %s",
            (mall_id,),
        )
        row = cur.fetchone()
    if not row:
        return None
    return {"access_token": row[0], "refresh_token": row[1], "expires_at": row[2]}


def _cafe24_token_request(form):
    """POST /oauth/token (Basic 인증). 응답 토큰을 DB에 저장하고 반환."""
    mall_id, client_id, client_secret, _ = _cafe24_cfg()
    url = f"https://{mall_id}.cafe24api.com/api/v2/oauth/token"
    basic = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()
    req = urllib.request.Request(
        url,
        data=urllib.parse.urlencode(form).encode(),
        method="POST",
        headers={
            "Authorization": f"Basic {basic}",
            "Content-Type": "application/x-www-form-urlencoded",
        },
    )
    with urllib.request.urlopen(req, timeout=25) as resp:
        tok = json.loads(resp.read().decode("utf-8"))
    _cafe24_save_token(mall_id, tok)
    return tok


def cafe24_exchange_code(code):
    """3~4단계: 코드 → 액세스/갱신 토큰 교환."""
    _, _, _, redirect_uri = _cafe24_cfg()
    return _cafe24_token_request({
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": redirect_uri,
    })


def cafe24_refresh_token():
    """갱신 토큰으로 새 액세스 토큰 발급(자동 갱신용)."""
    tok = _cafe24_load_token()
    if not tok or not tok.get("refresh_token"):
        raise RuntimeError("갱신할 refresh_token 이 없습니다. 다시 연결(승인)하세요.")
    return _cafe24_token_request({
        "grant_type": "refresh_token",
        "refresh_token": tok["refresh_token"],
    })


def cafe24_get(path, params=None):
    """액세스 토큰으로 GET. 401(만료)이면 자동 갱신 후 1회 재시도."""
    mall_id = get_secret("cafe24_mall_id")
    tok = _cafe24_load_token()
    if not tok:
        raise RuntimeError("카페24가 아직 연결되지 않았습니다. 먼저 연결(승인)하세요.")

    def _call(access_token):
        url = f"https://{mall_id}.cafe24api.com{path}"
        if params:
            url += "?" + urllib.parse.urlencode(params)
        req = urllib.request.Request(url, headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
            "X-Cafe24-Api-Version": CAFE24_API_VERSION,
        })
        with urllib.request.urlopen(req, timeout=25) as resp:
            return json.loads(resp.read().decode("utf-8"))

    try:
        return _call(tok["access_token"])
    except urllib.error.HTTPError as e:
        if e.code == 401:  # 액세스 토큰 만료 → 자동 갱신 후 재시도
            newtok = cafe24_refresh_token()
            return _call(newtok["access_token"])
        raise


def cafe24_today_summary():
    """오늘 주문 건수와 매출(결제금액 합계). 금액·건수만 요청 — 고객정보 미수집."""
    today = datetime.date.today().strftime("%Y-%m-%d")
    cnt = cafe24_get(
        "/api/v2/admin/orders/count",
        {"start_date": today, "end_date": today},
    ).get("count", 0)

    total = 0.0
    offset, limit = 0, 100
    while True:
        data = cafe24_get("/api/v2/admin/orders", {
            "start_date": today, "end_date": today,
            "limit": limit, "offset": offset,
            # 금액 필드만 요청 — 이름/연락처/주소는 애초에 받지 않음
            "fields": "order_id,actual_order_amount,payment_amount",
        })
        orders = data.get("orders") or []
        if not orders:
            break
        for o in orders:
            amt = o.get("actual_order_amount") or o.get("payment_amount") or 0
            try:
                total += float(amt)
            except (TypeError, ValueError):
                pass
        if len(orders) < limit:
            break
        offset += limit
    return cnt, total


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

    # 노션으로 보내기 — 오늘 모은 날씨를 노션 DB에 한 줄씩 추가
    if get_secret("notion_token"):
        st.divider()
        st.caption("노션으로 보내기 ('테스트' 칸=날씨 내용, '담당자' 칸=지정한 사용자)")
        if st.button("📤 오늘 모은 서울 날씨를 노션에 한 줄씩 추가"):
            try:
                today_rows = fetch_weather_collected_today()
            except Exception as e:
                st.error(f"조회 실패: {e}")
                today_rows = []
            if not today_rows:
                st.info("오늘 수집한 날씨가 없습니다. 위에서 먼저 수집하세요.")
            else:
                rows = [
                    (r[0].strftime("%Y-%m-%d"), r[1], r[2], r[3]) for r in today_rows
                ]
                try:
                    n = add_weather_rows_to_notion(rows)
                    st.success(f"노션에 {n}줄 추가했습니다.")
                except Exception as e:
                    st.error(f"노션 추가 실패: {e}")


def render_cafe24():
    st.write(
        "카페24(몰: **ergo2**) 관리자 API로 **오늘 매출·주문 건수**만 봅니다. "
        "고객 이름·연락처·주소는 가져오지 않습니다(scope: 주문 읽기, 금액·건수만 집계)."
    )
    mall_id, client_id, client_secret, redirect_uri = _cafe24_cfg()
    if not (mall_id and client_id and client_secret and redirect_uri):
        st.warning(
            "secrets 에 cafe24_client_id / cafe24_client_secret / "
            "cafe24_redirect_uri 를 넣어주세요. (몰 아이디는 ergo2로 설정됨)"
        )
        return

    # OAuth 콜백: 승인 후 ?code= 로 돌아오면 토큰 교환
    code = st.query_params.get("code")
    if code:
        try:
            cafe24_exchange_code(code)
            st.query_params.clear()
            st.success("카페24 연결 완료! 토큰을 저장했습니다.")
        except Exception as e:
            st.error(f"토큰 교환 실패: {e}")

    try:
        tok = _cafe24_load_token()
    except Exception as e:
        st.error(f"토큰 조회 실패: {e}")
        return

    if not tok:
        st.info("아직 연결 전입니다. 아래 버튼으로 승인하세요.")
        st.link_button("🔗 카페24 연결(승인)", cafe24_authorize_url())
        return

    st.success("연결됨 · 요청 시 토큰이 만료됐으면 자동으로 갱신됩니다.")
    if st.button("📊 오늘 매출·주문 건수 보기", type="primary"):
        try:
            cnt, total = cafe24_today_summary()
            c1, c2 = st.columns(2)
            c1.metric("오늘 주문 건수", f"{cnt:,} 건")
            c2.metric("오늘 매출(결제금액 합계)", f"{int(total):,} 원")
        except Exception as e:
            st.error(f"조회 실패: {e}")
    if st.button("🔄 토큰 수동 갱신"):
        try:
            cafe24_refresh_token()
            st.success("토큰을 갱신했습니다.")
        except Exception as e:
            st.error(f"갱신 실패: {e}")


def render_admin():
    st.title("🛠️ 관리자")
    tab_u, tab_e = st.tabs(["사용자 관리", "접근 허용 이메일"])

    with tab_u:
        st.subheader("새 로컬 계정 만들기")
        with st.form("new_local_user"):
            nm = st.text_input("이름")
            uid = st.text_input("아이디(비우면 이름으로 자동생성)")
            if st.form_submit_button("계정 생성") and nm.strip():
                try:
                    uname, temp = db_create_local_user(nm.strip(), uid)
                    st.success(
                        f"생성됨 → 아이디 `{uname}` · 임시비밀번호 `{temp}` "
                        "(첫 로그인 시 변경 필요)"
                    )
                except Exception as e:
                    st.error(f"생성 실패(아이디 중복?): {e}")

        st.divider()
        st.subheader("사용자 목록 · 화면 권한")
        try:
            users = db_list_users()
        except Exception as e:
            st.error(f"목록 조회 실패: {e}")
            users = []
        for u in users:
            who = u["name"] or u["username"] or u["email"] or f"#{u['id']}"
            tag = " · 관리자" if u["is_admin"] else ("" if u["is_active"] else " · 비활성")
            with st.expander(f"{who} ({u['auth_type']}){tag}"):
                if u["is_admin"]:
                    st.caption("관리자는 모든 화면을 봅니다(고정).")
                cur_screens = u["screens"] or []
                chosen = []
                cols = st.columns(len(SCREENS))
                for i, (key, title, _fn) in enumerate(SCREENS):
                    checked = cols[i].checkbox(
                        title.split(" ", 1)[-1],
                        value=(key in cur_screens),
                        key=f"scr_{u['id']}_{key}",
                        disabled=u["is_admin"],
                    )
                    if checked:
                        chosen.append(key)
                c1, c2, c3 = st.columns(3)
                if c1.button("권한 저장", key=f"save_{u['id']}", disabled=u["is_admin"]):
                    db_update_user_screens(u["id"], chosen)
                    st.success("저장됨")
                    st.rerun()
                if c2.button(("비활성화" if u["is_active"] else "활성화"), key=f"act_{u['id']}"):
                    db_set_active(u["id"], not u["is_active"])
                    st.rerun()
                if u["auth_type"] == "local" and c3.button("임시비번 재발급", key=f"rst_{u['id']}"):
                    st.info(f"새 임시비밀번호: `{db_reset_password(u['id'])}`")

    with tab_e:
        st.subheader("접근 허용 이메일(도메인 예외)")
        st.caption(f"@{allowed_domain()} 이 아니어도 여기 있으면 구글 로그인 허용.")
        with st.form("add_allowed"):
            e = st.text_input("이메일 추가")
            if st.form_submit_button("추가") and e.strip():
                db_add_allowed_email(e.strip())
                st.success("추가됨")
                st.rerun()
        try:
            for em in db_list_allowed_emails():
                c1, c2 = st.columns([4, 1])
                c1.write(em)
                if c2.button("삭제", key=f"del_{em}"):
                    db_remove_allowed_email(em)
                    st.rerun()
        except Exception as e:
            st.error(f"조회 실패: {e}")


# 화면 등록표: (키, 메뉴제목, 렌더함수) — 여기 없으면 어디에도 안 보임
SCREENS = [
    ("search", "🔎 검색어 수집", render_search),
    ("orders", "📦 주문 수집", render_orders),
    ("weather", "🌤️ 날씨", render_weather),
    ("cafe24", "🛒 카페24", render_cafe24),
]
SCREEN_KEYS = [s[0] for s in SCREENS]


def main():
    # DB 준비 (검색결과·주문·날씨·카페24·사용자·허용이메일 테이블)
    try:
        init_db()
    except Exception as e:
        st.error(f"DB 연결 실패: {e}")
        st.stop()

    user = st.session_state.get("user")
    if not user:
        render_login()
        st.stop()

    if user.get("kind") == "local" and user.get("must_change_pw"):
        render_change_pw()
        st.stop()

    with st.sidebar:
        st.title("🔎 더에르고 검색기")
        st.caption(f"👤 {user.get('name')}" + (" (관리자)" if user.get("is_admin") else ""))
        if st.button("로그아웃"):
            st.session_state.pop("user", None)
            st.rerun()

    # 권한 있는 화면만 메뉴에 등록 → 없는 화면은 보이지도, 주소로도 접근 불가
    allowed = SCREEN_KEYS if user.get("is_admin") else (user.get("screens") or [])
    pages = [
        st.Page(fn, title=title, url_path=key)
        for key, title, fn in SCREENS
        if key in allowed
    ]
    if user.get("is_admin"):
        pages.append(st.Page(render_admin, title="🛠️ 관리자", url_path="admin"))

    if not pages:
        st.title("🔎 더에르고 검색기")
        st.info("접근 가능한 화면이 없습니다. 관리자에게 권한을 요청하세요.")
        st.stop()

    st.navigation(pages).run()


main()
