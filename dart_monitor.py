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

# 제11-3-16조(중대재해 발생사실) 서식의 주요 항목. 표에서 라벨 매칭용.
SUMMARY_FIELDS = [
    ("발생일시", ["발생일시", "발생 일시", "사고일시", "사고 일시"]),
    ("발생장소", ["발생장소", "발생 장소", "사고장소"]),
    ("사고내용", ["사고내용", "사고 내용", "재해내용", "발생경위", "재해 내용"]),
    ("피해규모", ["피해규모", "피해 규모", "인명피해", "사망", "부상"]),
    ("사고원인", ["사고원인", "사고 원인", "재해원인", "발생원인"]),
    ("향후대책", ["향후대책", "향후 대책", "재발방지", "조치사항", "조치 사항"]),
    ("회사영향", ["회사에 미치는 영향", "영향"]),
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

    @property
    def url(self) -> str:
        return VIEWER_URL.format(rcept_no=self.rcept_no)


def env(name: str, default: str | None = None, required: bool = False) -> str:
    v = os.environ.get(name, default)
    if required and not v:
        sys.exit(f"환경변수 {name} 가 설정되지 않았습니다.")
    return v or ""


# ---------------------------------------------------------------- DART API


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
        r = requests.get(LIST_URL, params=params, timeout=30)
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
    r = requests.get(DOC_URL, params={"crtfc_key": api_key, "rcept_no": rcept_no}, timeout=60)
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
            label = " ".join(row[:-1])
            value = row[-1]
            for key, aliases in SUMMARY_FIELDS:
                if key in summary:
                    continue
                if any(a in label for a in aliases) and value != label:
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
    details summary{cursor:pointer;color:#357;margin-top:8px}
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
        if d.error:
            parts.append(f"<div class='err'>{esc(d.error)}</div>")
        if d.summary:
            parts.append("<table>")
            for key, _ in SUMMARY_FIELDS:
                if key in d.summary:
                    parts.append(f"<tr><th>{esc(key)}</th><td>{esc(d.summary[key])}</td></tr>")
            parts.append("</table>")
        if d.tables:
            parts.append("<details><summary>공시 표 전체 보기</summary>")
            for rows in d.tables:
                parts.append("<table>")
                for row in rows:
                    parts.append("<tr>" + "".join(f"<td>{esc(c)}</td>" for c in row) + "</tr>")
                parts.append("</table>")
            parts.append("</details>")
        parts.append("</div>")
    parts.append("<div class='meta'>본 메일은 OpenDART API 기반 자동 발송입니다.</div>")
    parts.append("</body></html>")
    return "".join(parts)


def render_text(items: list[Disclosure]) -> str:
    lines = ["DART 중대재해 발생사실 공시 알림", ""]
    for d in items:
        lines.append(f"[{d.rcept_dt}] {d.corp_name} - {d.report_nm}")
        for key, _ in SUMMARY_FIELDS:
            if key in d.summary:
                lines.append(f"  {key}: {d.summary[key]}")
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


def load_state(path: Path) -> set[str]:
    if path.exists():
        try:
            return set(json.loads(path.read_text(encoding="utf-8")).get("sent", []))
        except json.JSONDecodeError:
            return set()
    return set()


def save_state(path: Path, sent: set[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    # 최근 2000건만 유지
    keep = sorted(sent)[-2000:]
    path.write_text(json.dumps({"sent": keep}, ensure_ascii=False, indent=0), encoding="utf-8")


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
    all_items = fetch_list(api_key, bgn_de, end_de)
    print(f"전체 공시 {len(all_items)}건")

    matched = [
        it for it in all_items if any(k in (it.get("report_nm") or "") for k in keywords)
    ]
    print(f"키워드 일치 {len(matched)}건")

    sent = load_state(state_path)
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
            report_nm=it.get("report_nm", ""),
            flr_nm=it.get("flr_nm", ""),
        )
        enrich(api_key, d)
        disclosures.append(d)
        print(f"  - {d.rcept_dt} {d.corp_name} | {d.report_nm} | {d.url}")
        if d.error:
            print(f"    ! {d.error}")

    subject = f"[DART 중대재해] {now:%m/%d} 신규 {len(disclosures)}건 - " + ", ".join(
        dict.fromkeys(d.corp_name for d in disclosures)
    )
    if len(subject) > 120:
        subject = subject[:117] + "..."

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
        sent.update(d.rcept_no for d in disclosures)
        save_state(state_path, sent)
    return 0


if __name__ == "__main__":
    sys.exit(main())
