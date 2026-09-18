# DART 중대재해 발생사실 공시 모니터링

매일 15:00(KST) OpenDART API로 전 기업 공시를 확인하여, 보고서명에 "중대재해"가 포함된
신규 공시가 있으면 사고내역(발생일시·장소·내용·피해규모·원인·향후대책)을 추출해 HTML 메일로 발송합니다.
신규 건이 없으면 메일을 보내지 않습니다.

## 동작 방식

1. `list.json`(공시검색)으로 최근 3일 공시 목록을 전부 조회 (주말·지연 대비)
2. 보고서명 키워드 필터 → `state/sent.json`에 없는 건만 신규로 판단
3. `document.xml`(원본파일)로 공시 원문을 받아 표를 파싱, 제11-3-16조 항목 라벨로 요약 추출
4. Gmail SMTP로 발송 후 발송 이력을 저장소에 커밋

## 설정 (GitHub Actions)

1. 이 폴더를 GitHub 저장소(private 권장)로 푸시
2. **Settings → Secrets and variables → Actions → Secrets** 에 등록
   | 이름 | 값 |
   |---|---|
   | `DART_API_KEY` | OpenDART 인증키 (https://opendart.fss.or.kr 가입 후 발급, 즉시 발급됨) |
   | `SMTP_USER` | 발신 Gmail 주소 |
   | `SMTP_PASS` | Gmail **앱 비밀번호** (Google 계정 → 보안 → 2단계 인증 → 앱 비밀번호) |
   | `MAIL_TO` | 수신자, 쉼표 구분 |
3. (선택) **Variables** 에 `KEYWORDS` 등록으로 필터 키워드 변경 (예: `중대재해,산업재해`)
4. **Actions** 탭 → "DART 중대재해 공시 모니터링" → **Run workflow** 에서
   `dry_run=1`, `lookback_days=90` 으로 한번 실행하여 과거 건이 잘 추출되는지 확인
   (Artifacts의 `preview.html`에서 메일 미리보기 확인)
5. 정상 확인 후 `dry_run=0`으로 실행하면 실제 발송. 이후 매 평일 15:00 자동 실행

## 로컬 테스트

```bash
pip install -r requirements.txt
copy .env.example .env   # 값 채우기
# PowerShell
Get-Content .env | ForEach-Object { if ($_ -match '^(\w+)=(.*)$') { $env:($matches[1]) = $matches[2] } }
$env:LOOKBACK_DAYS="90"; python dart_monitor.py
```

## 주의

- GitHub cron은 정각보다 수 분~수십 분 늦게 실행될 수 있습니다.
- OpenDART 인증키는 일 10,000회 호출 제한이 있습니다. 하루 실행에 필요한 호출은 20~40회 수준입니다.
- 거래소 공시(KIND) 원문 구조에 따라 표 추출이 안 되는 경우 원문 링크와 함께 오류 메시지를 표기합니다.
