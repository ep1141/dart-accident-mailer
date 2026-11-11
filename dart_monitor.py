"""
DART 중대재해 발생사실 공시 모니터링 & 메일 발송

- OpenDART 공시검색 API로 최근 N일 전 기업 공시 목록을 조회
- 보고서명에 키워드(기본: 중대재해)가 포함된 공시만 필터
- 공시 원본(document.xml)을 받아 표(사고내용/피해규모/원인/대책 등)를 추출
- 신규 건이 있으면 HTML 메일 발송, 발송 이력은 state/sent.json 에 저장(중복 방지)

필요 환경변수
  DART_API_KEY   OpenDART 인증키
  SMTP_HOST      기본 smtp.gmail.com
  SMTP_PORT      기본 587
  SMTP_USER      발신 계정 (Gmail 주소)
  SMTP_PASS      Gmail 앱 비밀번호
  MAIL_TO        수신자 (쉼표 구분)
  MAIL_FROM      발신자 표기 (기본 SMTP_USER)
선택
  KEYWORDS       보고서명 필터 키워드, 쉼표 구분 (기본 "중대재해")
  LOOKBACK_DAYS  조회 기간(일), 기본 3 (주말/지연 대비)
  DRY_RUN        "1"이면 메일을 보내지 않고 결과만 출력
  STATE_FILE     기본 state/sent.json
"""

from __future__ import annotations

import html
import io
import json
import os
import re
import smtplib
import sys
import time
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import requests
from bs4 import BeautifulSoup

KST = timezone(timedelta(hours=9))
LIST_URL = "https://opendart.fss.or.kr/api/list.json"
DOC_URL = "https://opendart.fss.or.kr/api/document.xml"
VIEWER_URL = "https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcept_no}"

# 제11-3-16조(중대재해 발생사실) 서식 항목. (표시명, 표 라벨 셀에 포함되는 문자열들)
# 실제 공시 표 구조: 1. 중대재해내용 > 발생 장소 / 발생 재해 내용 / 사망자 수 / 부상자 수,
# 2. 중대재해 발생일자, 3. 고용노동부 보고일자, 4. 조치사항 및 향후대책, 5. 기타 투자판단과 관련한 중요사항
SUMMARY_FIELDS = [
    ("종속회사명", ["종속회사명", "자회사명"]),
    ("정정사유", ["정정사유"]),
    ("발생 장소", ["발생 장소", "발생장소", "사고장소"]),
    ("발생 재해 내용", ["발생 재해 내용", "발생재해내용", "사고내용", "사고 내용", "재해 내용"]),
    ("사망자 수", ["사망자 수", "사망자수"]),
    ("부상자 수", ["부상자 수", "부상자수"]),
    ("발생일자", ["발생일자", "발생 일자", "발생일시", "발생 일시"]),
    ("고용노동부 보고일자", ["고용노동부 보고일자", "보고일자"]),
    ("조치사항 및 향후대책", ["조치사항 및 향후대책", "향후대책", "향후 대책", "조치사항"]),
    ("기타 중요사항", ["기타 투자판단", "기타 중요사항"]),
]

CORP_CLS = {"Y": "유가", "K": "코스닥", "N": "코넥스", "E": "기타"}


@dataclass
class Disclosure:
    rcept_no: str
    rcept_dt: str
    corp_name: str
    corp_cls: str
    stock_code: str
    report_nm: str
    flr_nm: str
    summary: dict = field(default_factory=dict)
    tables: list = field(default_factory=list)  # list[list[list[str]]]
    error: str | None = None
    also_filed: list = field(default_factory=list)  # 같은 사고를 공시한 다른 회사(모회사 등)
    en: dict = field(default_factory=dict)  # 영문 번역 {"corp_name","report_nm","fields":{한글라벨: 영문값}}

    @property
    def url(self) -> str:
        return VIEWER_URL.format(rcept_no=self.rcept_no)

    @property
    def is_parent_filing(self) -> bool:
        """모회사가 자회사·종속회사 사고를 대신 공시한 건인지."""
        return any(k in self.report_nm for k in ("종속회사", "자회사"))

    @property
    def is_correction(self) -> bool:
        return "정정" in self.report_nm

    def content_key(self) -> str | None:
        """같은 사고를 식별하는 키. 재해 내용이 없으면 None(중복 판단 안 함)."""
        content = self.summary.get("발생 재해 내용") or ""
        if not content:
            return None
        norm = re.sub(r"[\s\W_]+", "", content)[:80]
        date = re.sub(r"\D", "", self.summary.get("발생일자") or "")
        return f"{'C' if self.is_correction else 'O'}|{date}|{norm}"


def dedupe_same_incident(items: list[Disclosure], sent_keys: set[str]) -> list[Disclosure]:
    """자회사와 모회사가 같은 사고를 각각 공시한 경우 하나만 남긴다.
    자회사 직접 공시를 우선하고, 나머지 공시 회사는 also_filed에 기록.
    이미 발송한 사고(sent_keys)와 같은 건이면 제외."""
    groups: dict[str, list[Disclosure]] = {}
    kept: list[Disclosure] = []
    for d in items:
        key = d.content_key()
        if key is None:
            kept.append(d)
        elif key in sent_keys:
            print(f"    (중복 제외: 이미 발송한 사고) {d.corp_name} | {d.report_nm}")
        else:
            groups.setdefault(key, []).append(d)
    for ds in groups.values():
        ds.sort(key=lambda x: (x.is_parent_filing, x.rcept_no))
        primary = ds[0]
        for other in ds[1:]:
            primary.also_filed.append(other)
            print(f"    (중복 병합) {other.corp_name} | {other.report_nm} -> {primary.corp_name}")
        kept.append(primary)
    kept.sort(key=lambda x: (x.rcept_dt, x.rcept_no))
    return kept


def env(name: str, default: str | None = None, required: bool = False) -> str:
    v = os.environ.get(name, default)
    if required and not v:
        sys.exit(f"환경변수 {name} 가 설정되지 않았습니다.")
    return v or ""


# ---------------------------------------------------------------- DART API


def dart_get(url: str, params: dict, read_timeout: int = 60, attempts: int = 5) -> requests.Response:
    """DART는 해외 클라우드(GitHub 러너)에서 접속이 느리거나 끊기는 경우가 있어 재시도한다."""
    last: Exception | None = None
    for i in range(1, attempts + 1):
        try:
            r = requests.get(url, params=params, timeout=(20, read_timeout))
            if r.status_code >= 500:
                raise requests.HTTPError(f"HTTP {r.status_code}")
            return r
        except (requests.ConnectionError, requests.Timeout, requests.HTTPError) as e:
            last = e
            wait = min(60, 5 * 2 ** (i - 1))
            print(f"    DART 접속 재시도 {i}/{attempts} ({type(e).__name__}) - {wait}초 후")
            if i < attempts:
                time.sleep(wait)
    raise last  # type: ignore[misc]


def fetch_list(api_key: str, bgn_de: str, end_de: str) -> list[dict]:
    """기간 내 전체 공시 목록(전 기업, 전 유형)을 페이지 순회하며 수집."""
    items: list[dict] = []
    page = 1
    while True:
        params = {
            "crtfc_key": api_key,
            "bgn_de": bgn_de,
            "end_de": end_de,
            "page_no": page,
            "page_count": 100,
        }
        r = dart_get(LIST_URL, params)
        r.raise_for_status()
        data = r.json()
        status = data.get("status")
        if status == "013":  # 조회된 데이터가 없습니다
            break
        if status != "000":
            raise RuntimeError(f"DART list API 오류 {status}: {data.get('message')}")
        items.extend(data.get("list", []))
        total_page = int(data.get("total_page", 1))
        if page >= total_page:
            break
        page += 1
    return items


def fetch_document_text(api_key: str, rcept_no: str) -> str:
    """공시 원본 zip을 받아 XML/HTML 본문 문자열로 반환."""
    r = dart_get(DOC_URL, {"crtfc_key": api_key, "rcept_no": rcept_no})
    r.raise_for_status()
    ctype = r.headers.get("Content-Type", "")
    if "json" in ctype or r.content[:1] == b"{":
        data = r.json()
        raise RuntimeError(f"DART document API 오류 {data.get('status')}: {data.get('message')}")
    zf = zipfile.ZipFile(io.BytesIO(r.content))
    names = [n for n in zf.namelist() if n.lower().endswith((".xml", ".html", ".htm"))]
    if not names:
        raise RuntimeError("원본 zip 안에 문서 파일이 없습니다.")
    # 파일이 여러 개면 가장 큰 본문 사용
    name = max(names, key=lambda n: zf.getinfo(n).file_size)
    raw = zf.read(name)
    for enc in ("utf-8", "euc-kr", "cp949"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


# ---------------------------------------------------------------- 파싱


def _cell_text(tag) -> str:
    text = tag.get_text(" ", strip=True)
    return re.sub(r"\s+", " ", text).strip()


def parse_tables(doc: str) -> list[list[list[str]]]:
    """DART XML(TABLE/TR/TD/TE/TH) 또는 HTML 표를 2차원 문자열 리스트로 변환."""
    soup = BeautifulSoup(doc, "html.parser")
    tables = []
    for t in soup.find_all(re.compile(r"^table$", re.I)):
        rows = []
        for tr in t.find_all(re.compile(r"^tr$", re.I)):
            cells = [
                _cell_text(td)
                for td in tr.find_all(re.compile(r"^(td|te|th|tu)$", re.I))
            ]
            cells = [c for c in cells if c != ""]
            if cells:
                rows.append(cells)
        if rows:
            tables.append(rows)
    return tables


def summarize(tables: list[list[list[str]]]) -> dict:
    """표의 라벨 셀을 기준으로 주요 항목을 추출."""
    summary: dict[str, str] = {}
    for rows in tables:
        for row in rows:
            if len(row) < 2:
                continue
            # 라벨 셀 바로 오른쪽 셀을 값으로 사용
            # 예) ["1. 중대재해내용", "발생 장소", "온산제련소 2공장"] -> 발생 장소 = 온산제련소 2공장
            #     ["종속회사명", "한화솔루션㈜", "영문", "HANWHA ..."] -> 종속회사명 = 한화솔루션㈜
            #     정정공시 ["4. 조치사항 및 향후대책", "정정전 값", "정정후 값"] -> 마지막 셀(정정후)
            for i, cell in enumerate(row[:-1]):
                for key, aliases in SUMMARY_FIELDS:
                    if key in summary:
                        continue
                    if any(a in cell for a in aliases):
                        value = row[-1] if len(row) - i == 3 and i == 0 else row[i + 1]
                        if value and value != cell:
                            summary[key] = value
                        break
    return summary


def enrich(api_key: str, d: Disclosure) -> None:
    try:
        doc = fetch_document_text(api_key, d.rcept_no)
        d.tables = parse_tables(doc)
        d.summary = summarize(d.tables)
        if not d.tables:
            d.error = "표를 찾지 못했습니다. 원문 링크를 확인하세요."
    except Exception as e:  # noqa: BLE001
        d.error = f"원문 추출 실패: {e}"


# ---------------------------------------------------------------- 영문 번역

EN_LABELS = {
    "종속회사명": "Subsidiary",
    "정정사유": "Reason for correction",
    "발생 장소": "Location",
    "발생 재해 내용": "Description of accident",
    "사망자 수": "Fatalities",
    "부상자 수": "Injuries",
    "발생일자": "Date of occurrence",
    "고용노동부 보고일자": "Reported to Ministry of Employment and Labor",
    "조치사항 및 향후대책": "Actions taken and future measures",
    "기타 중요사항": "Other important matters",
}

# 번역 방식(위에서부터 사용 가능한 것을 순서대로 시도, 모두 실패하면 한글만 발송)
#  1) Anthropic API    : ANTHROPIC_API_KEY 가 있을 때만 (유료, 선택)
#  2) GitHub Models    : Actions의 GITHUB_TOKEN 으로 무료 사용 (워크플로우 permissions: models: read)
#  3) Google 무료 엔드포인트 : 키 불필요, 비공식이라 최후 수단
GH_MODELS_URL = "https://models.github.ai/inference/chat/completions"
GH_MODEL = "openai/gpt-4o-mini"
ANTHROPIC_MODEL = "claude-haiku-4-5-20251001"


def _llm_prompt(payload: dict) -> str:
    return (
        "Translate the values of this Korean corporate disclosure about a serious industrial accident "
        "into clear, neutral English. Keep keys unchanged, keep numbers and dates as-is, "
        "romanize or use the official English name for company names, and do not add or omit facts. "
        "Return only the JSON object.\n\n" + json.dumps(payload, ensure_ascii=False)
    )


def _parse_json_object(text: str) -> dict:
    m = re.search(r"\{.*\}", text, re.S)
    if not m:
        raise ValueError("번역 응답에서 JSON을 찾지 못함")
    return json.loads(m.group(0))


def _translate_anthropic(payload: dict) -> dict:
    r = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": env("ANTHROPIC_API_KEY"),
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        json={
            "model": env("TRANSLATE_MODEL", ANTHROPIC_MODEL),
            "max_tokens": 2000,
            "messages": [{"role": "user", "content": _llm_prompt(payload)}],
        },
        timeout=60,
    )
    r.raise_for_status()
    return _parse_json_object("".join(b.get("text", "") for b in r.json().get("content", [])))


def _translate_github_models(payload: dict) -> dict:
    r = requests.post(
        GH_MODELS_URL,
        headers={
            "Authorization": f"Bearer {env('GITHUB_TOKEN')}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "Content-Type": "application/json",
        },
        json={
            "model": env("GH_MODEL", GH_MODEL),
            "messages": [{"role": "user", "content": _llm_prompt(payload)}],
            "temperature": 0,
        },
        timeout=90,
    )
    r.raise_for_status()
    return _parse_json_object(r.json()["choices"][0]["message"]["content"])


def _gtx(text: str) -> str:
    """Google 무료 번역 엔드포인트(키 불필요). 긴 글은 1,500자 단위로 나눠 호출."""
    chunks, cur = [], ""
    for piece in re.split(r"(?<=[.!?\n])\s*", text):
        if len(cur) + len(piece) > 1500 and cur:
            chunks.append(cur)
            cur = ""
        cur += piece + " "
    if cur.strip():
        chunks.append(cur.strip())
    out = []
    for c in chunks:
        r = requests.get(
            "https://translate.googleapis.com/translate_a/single",
            params={"client": "gtx", "sl": "ko", "tl": "en", "dt": "t", "q": c},
            timeout=30,
        )
        r.raise_for_status()
        out.append("".join(seg[0] for seg in r.json()[0] if seg and seg[0]))
    return " ".join(out).strip()


def _translate_google(payload: dict) -> dict:
    def tr(v: str) -> str:
        return _gtx(v) if re.search(r"[가-힣]", v or "") else v

    return {
        "corp_name": tr(payload["corp_name"]),
        "report_nm": tr(payload["report_nm"]),
        "fields": {k: tr(v) for k, v in payload["fields"].items()},
    }


def translate_disclosure(d: Disclosure) -> None:
    """공시 요약을 영문으로 번역해 d.en에 저장. 모두 실패하면 건너뛴다(한글 메일만 발송)."""
    if not d.summary:
        return
    payload = {
        "corp_name": d.corp_name,
        "report_nm": d.report_nm,
        "fields": {k: v for k, v in d.summary.items() if k in EN_LABELS},
    }
    providers = []
    if env("ANTHROPIC_API_KEY"):
        providers.append(("Anthropic", _translate_anthropic))
    if env("GITHUB_TOKEN"):
        providers.append(("GitHub Models", _translate_github_models))
    providers.append(("Google 무료", _translate_google))
    for name, fn in providers:
        try:
            data = fn(payload)
            if isinstance(data.get("fields"), dict) and data["fields"]:
                d.en = data
                print(f"    영문 번역 완료({name}): {d.corp_name}")
                return
        except Exception as e:  # noqa: BLE001
            print(f"    ! 영문 번역 실패({name}, {d.corp_name}): {e}")


# ---------------------------------------------------------------- 메일


def esc(s: str) -> str:
    return html.escape(s or "")


def render_html(items: list[Disclosure], run_time: datetime) -> str:
    css = """
    body{font-family:'Malgun Gothic',Apple SD Gothic Neo,sans-serif;font-size:14px;color:#222}
    h2{margin:0 0 6px}
    .meta{color:#666;margin-bottom:18px}
    .card{border:1px solid #ddd;border-radius:6px;padding:14px 16px;margin-bottom:20px}
    .card h3{margin:0 0 8px;font-size:16px}
    .tag{display:inline-block;background:#eef;color:#335;border-radius:3px;padding:1px 6px;font-size:12px;margin-right:6px}
    table{border-collapse:collapse;width:100%;margin-top:8px}
    td,th{border:1px solid #ddd;padding:6px 8px;vertical-align:top;font-size:13px}
    th{background:#f5f5f5;text-align:left;width:22%}
    .err{color:#a00}
    .sub{color:#666;font-size:12px;margin-top:14px}
    .also{color:#666;font-size:12px;margin-top:4px}
    table.raw td{color:#555;font-size:12px}
    """
    parts = [f"<html><head><meta charset='utf-8'><style>{css}</style></head><body>"]
    parts.append("<h2>DART 중대재해 발생사실 공시 알림</h2>")
    parts.append(
        f"<div class='meta'>확인 시각: {run_time:%Y-%m-%d %H:%M} KST · 신규 {len(items)}건</div>"
    )
    for d in items:
        parts.append("<div class='card'>")
        parts.append(
            f"<h3>{esc(d.corp_name)} <span class='tag'>{esc(CORP_CLS.get(d.corp_cls, d.corp_cls))}</span>"
            f"<span class='tag'>{esc(d.rcept_dt)}</span></h3>"
        )
        parts.append(
            f"<div>{esc(d.report_nm)} · 제출인 {esc(d.flr_nm)} · "
            f"<a href='{d.url}'>DART 원문 보기</a></div>"
        )
        if d.also_filed:
            links = " · ".join(
                f"<a href='{o.url}'>{esc(o.corp_name)}</a> ({esc(o.report_nm)})" for o in d.also_filed
            )
            parts.append(f"<div class='also'>같은 사고 공시: {links}</div>")
        if d.error:
            parts.append(f"<div class='err'>{esc(d.error)}</div>")
        if d.summary:
            parts.append("<table>")
            for key, _ in SUMMARY_FIELDS:
                if key in d.summary:
                    parts.append(f"<tr><th>{esc(key)}</th><td>{esc(d.summary[key])}</td></tr>")
            parts.append("</table>")
        # 요약 표를 뽑지 못한 경우에만 원문 표를 대신 보여준다
        if d.tables and not d.summary:
            parts.append("<div class='sub'>요약 추출 실패 · 공시 원문 표</div>")
            for rows in d.tables:
                parts.append("<table class='raw'>")
                for row in rows:
                    parts.append("<tr>" + "".join(f"<td>{esc(c)}</td>" for c in row) + "</tr>")
                parts.append("</table>")
        parts.append("</div>")
    parts.append("<div class='meta'>본 메일은 OpenDART API 기반 자동 발송입니다.</div>")

    # 영문 번역본 (번역에 성공한 건만)
    en_items = [d for d in items if d.en]
    if en_items:
        parts.append("<hr style='border:0;border-top:2px solid #999;margin:28px 0 18px'>")
        parts.append("<h2>Serious Accident Disclosure Alert (English)</h2>")
        parts.append("<div class='meta'>Machine-translated from the Korean disclosure. "
                     "Please refer to the original DART filing for accuracy.</div>")
        for d in en_items:
            parts.append("<div class='card'>")
            parts.append(
                f"<h3>{esc(d.en.get('corp_name') or d.corp_name)} "
                f"<span class='tag'>{esc(d.rcept_dt)}</span></h3>"
            )
            parts.append(
                f"<div>{esc(d.en.get('report_nm') or d.report_nm)} · "
                f"<a href='{d.url}'>View original on DART</a></div>"
            )
            parts.append("<table>")
            for key, _ in SUMMARY_FIELDS:
                if key in d.en["fields"]:
                    parts.append(
                        f"<tr><th>{esc(EN_LABELS[key])}</th><td>{esc(str(d.en['fields'][key]))}</td></tr>"
                    )
            parts.append("</table></div>")
    parts.append("</body></html>")
    return "".join(parts)


def render_text(items: list[Disclosure]) -> str:
    lines = ["DART 중대재해 발생사실 공시 알림", ""]
    for d in items:
        lines.append(f"[{d.rcept_dt}] {d.corp_name} - {d.report_nm}")
        if d.also_filed:
            lines.append("  같은 사고 공시: " + ", ".join(o.corp_name for o in d.also_filed))
        for key, _ in SUMMARY_FIELDS:
            if key in d.summary:
                lines.append(f"  {key}: {d.summary[key]}")
        lines.append(f"  {d.url}")
        lines.append("")
    en_items = [d for d in items if d.en]
    if en_items:
        lines += ["-" * 40, "Serious Accident Disclosure Alert (English, machine-translated)", ""]
        for d in en_items:
            lines.append(f"[{d.rcept_dt}] {d.en.get('corp_name') or d.corp_name} - {d.en.get('report_nm') or d.report_nm}")
            for key, _ in SUMMARY_FIELDS:
                if key in d.en["fields"]:
                    lines.append(f"  {EN_LABELS[key]}: {d.en['fields'][key]}")
            lines.append(f"  {d.url}")
            lines.append("")
    return "\n".join(lines)


def send_mail(subject: str, html_body: str, text_body: str) -> None:
    host = env("SMTP_HOST", "smtp.gmail.com")
    port = int(env("SMTP_PORT", "587"))
    user = env("SMTP_USER", required=True)
    password = env("SMTP_PASS", required=True)
    sender = env("MAIL_FROM", user)
    recipients = [x.strip() for x in env("MAIL_TO", required=True).split(",") if x.strip()]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg.attach(MIMEText(text_body, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    with smtplib.SMTP(host, port, timeout=60) as s:
        s.ehlo()
        s.starttls()
        s.login(user, password)
        s.sendmail(sender, recipients, msg.as_string())


# ---------------------------------------------------------------- 상태


def load_state(path: Path) -> tuple[set[str], set[str]]:
    """(발송한 공시번호, 발송한 사고 키) 반환."""
    if path.exists():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            return set(data.get("sent", [])), set(data.get("keys", []))
        except json.JSONDecodeError:
            pass
    return set(), set()


def save_state(path: Path, sent: set[str], keys: set[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    # 최근 2000건만 유지
    data = {"sent": sorted(sent)[-2000:], "keys": sorted(keys)[-2000:]}
    path.write_text(json.dumps(data, ensure_ascii=False, indent=0), encoding="utf-8")


# ---------------------------------------------------------------- main


def main() -> int:
    api_key = env("DART_API_KEY", required=True)
    keywords = [k.strip() for k in env("KEYWORDS", "중대재해").split(",") if k.strip()]
    lookback = int(env("LOOKBACK_DAYS", "3"))
    dry_run = env("DRY_RUN", "0") == "1"
    state_path = Path(env("STATE_FILE", "state/sent.json"))

    now = datetime.now(KST)
    end_de = now.strftime("%Y%m%d")
    bgn_de = (now - timedelta(days=lookback)).strftime("%Y%m%d")

    print(f"[{now:%Y-%m-%d %H:%M}] 공시 조회 {bgn_de}~{end_de}, 키워드 {keywords}")
    try:
        all_items = fetch_list(api_key, bgn_de, end_de)
    except (requests.ConnectionError, requests.Timeout) as e:
        # 예약 실행은 매시간 반복되므로, DART 접속 불가 시 실패 알림을 남발하지 않고 다음 실행에서 재시도한다.
        print(f"DART 접속 실패: {e}")
        if env("GITHUB_EVENT_NAME") == "schedule":
            print("예약 실행이므로 정상 종료하고 다음 실행에서 재시도합니다.")
            return 0
        return 1
    print(f"전체 공시 {len(all_items)}건")

    matched = [
        it for it in all_items if any(k in (it.get("report_nm") or "") for k in keywords)
    ]
    print(f"키워드 일치 {len(matched)}건")

    sent, sent_keys = load_state(state_path)
    new_items = [it for it in matched if it["rcept_no"] not in sent]
    print(f"신규(미발송) {len(new_items)}건")

    if not new_items:
        print("신규 건 없음. 메일을 보내지 않습니다.")
        return 0

    disclosures: list[Disclosure] = []
    for it in sorted(new_items, key=lambda x: (x["rcept_dt"], x["rcept_no"])):
        d = Disclosure(
            rcept_no=it["rcept_no"],
            rcept_dt=it.get("rcept_dt", ""),
            corp_name=it.get("corp_name", ""),
            corp_cls=it.get("corp_cls", ""),
            stock_code=it.get("stock_code", ""),
            report_nm=(it.get("report_nm") or "").strip(),
            flr_nm=it.get("flr_nm", ""),
        )
        enrich(api_key, d)
        disclosures.append(d)
        print(f"  - {d.rcept_dt} {d.corp_name} | {d.report_nm} | {d.url}")
        if d.error:
            print(f"    ! {d.error}")

    all_fetched = list(disclosures)
    disclosures = dedupe_same_incident(disclosures, sent_keys)
    print(f"중복 병합 후 {len(disclosures)}건")
    if not disclosures:
        print("발송할 신규 사고 없음(모두 기발송 사고의 중복 공시). 이력만 갱신합니다.")
        if not dry_run:
            sent.update(d.rcept_no for d in all_fetched)
            save_state(state_path, sent, sent_keys)
        return 0

    subject = f"[CSO 중대재해 공유] {now:%m/%d} " + ", ".join(
        dict.fromkeys(d.corp_name for d in disclosures)
    )
    if len(subject) > 120:
        subject = subject[:117] + "..."

    for d in disclosures:
        translate_disclosure(d)

    html_body = render_html(disclosures, now)
    text_body = render_text(disclosures)

    if dry_run:
        out = Path("out")
        out.mkdir(exist_ok=True)
        (out / "preview.html").write_text(html_body, encoding="utf-8")
        print(f"DRY_RUN: 메일 미발송. 미리보기 out/preview.html 저장. 제목: {subject}")
    else:
        send_mail(subject, html_body, text_body)
        print(f"메일 발송 완료: {subject}")
        sent.update(d.rcept_no for d in all_fetched)
        sent_keys.update(k for k in (d.content_key() for d in disclosures) if k)
        save_state(state_path, sent, sent_keys)
    return 0


if __name__ == "__main__":
    sys.exit(main())
