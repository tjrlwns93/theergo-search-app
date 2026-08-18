# -*- coding: utf-8 -*-
"""
더에르고 검색 웹앱 (Streamlit)
- 비밀번호 로그인 후에만 화면이 보인다.
- [검색어]를 넣으면 theergo.co.kr 첫 번째 결과의 제목/링크를 가져온다.
- 결과를 Supabase(PostgreSQL)에 저장하고, 저장 이력을 보여준다.
- 접속정보/비밀번호는 코드가 아니라 secrets(.streamlit/secrets.toml)에서 읽는다.
"""
import html as html_mod
import io
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
def main():
    st.title("🔎 더에르고 검색기")
    st.write("검색어를 넣으면 theergo.co.kr 첫 번째 결과의 제목·링크를 가져와 저장합니다.")

    # DB 준비
    try:
        init_db()
    except Exception as e:
        st.error(f"DB 연결 실패: {e}")
        st.stop()

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


if not check_password():
    st.stop()

main()
