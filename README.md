# 더에르고 검색기 (theergo search app)

[theergo.co.kr](https://theergo.co.kr) 에서 검색어별 첫 번째 결과의 **제목·링크**를 가져와
**Supabase(PostgreSQL)** 에 저장하는 Streamlit 웹앱. 비밀번호 로그인 후에만 화면이 보입니다.

## 로컬 실행

```bash
pip install -r requirements.txt
# 비밀정보 설정 (한 번만)
cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # 값 채우기
streamlit run app.py
```

`.streamlit/secrets.toml` 에 값을 넣습니다. (자세한 예시는 `secrets.toml.example`)

```toml
app_password = "로그인 비밀번호"

# Supabase 대시보드 > Connect > "Session pooler" 값을 그대로 복사
db_host = "aws-0-ap-northeast-2.pooler.supabase.com"
db_port = 5432
db_user = "postgres.<프로젝트ref>"
db_password = "실제 DB 비밀번호"
db_name = "postgres"
```

> `.streamlit/secrets.toml` 는 `.gitignore` 에 등록되어 **깃허브에 올라가지 않습니다.**

## 배포 (Streamlit Community Cloud)

1. 이 저장소를 연결하고 `app.py` 를 지정합니다.
2. **App settings → Secrets** 에 `app_password` 와 DB 접속정보(`db_host` 등)를 붙여넣습니다.
3. `requirements.txt` 를 보고 클라우드가 자동으로 패키지를 설치합니다.

> ⚠️ **연결 주소 주의:** Supabase 의 **direct** 주소(`db.xxxx.supabase.co`)는 IPv6 전용이라
> 일반 PC / Streamlit Cloud(IPv4)에서는 `could not translate host name ...` 오류가 납니다.
> 반드시 대시보드 **Connect → Session pooler** 의 값을 쓰세요. `db_host` 에는 거기 보이는
> host(예: `aws-0-ap-northeast-2.pooler.supabase.com`)를 **그대로** 붙여넣습니다.

## 파일

| 파일 | 설명 |
|---|---|
| `app.py` | Streamlit 웹앱 본체 |
| `theergo_search.py` | 명령줄(CLI) 버전 — 엑셀 입력/엑셀 출력 |
| `requirements.txt` | 설치할 패키지 목록 |
| `.streamlit/secrets.toml.example` | secrets 예시 (실제 값 X) |
