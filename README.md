# GitHub Actions용 온라인 가격 모니터링 패키지

## 포함 파일
- `price_monitor_step1_cloud.py` : 클라우드/GitHub Actions용 수정본
- `requirements.txt`
- `.github/workflows/daily-price-monitor.yml`
- `.env.example`

## 추가로 직접 넣어야 하는 파일
같은 폴더(저장소 루트)에 아래 파일을 같이 올려야 합니다.
- `펌프리스트_시트.xlsx`  **필수**
- `통합결과양식.xlsx`  필요 시
- `crawler_step2.py`  사용 시
- `매크로모음.bas`  self-hosted Windows runner에서 VBA 사용 시

## GitHub 저장소 구조 예시
```text
repo-root/
├─ price_monitor_step1_cloud.py
├─ requirements.txt
├─ 펌프리스트_시트.xlsx
├─ 통합결과양식.xlsx
├─ crawler_step2.py
├─ 매크로모음.bas
└─ .github/
   └─ workflows/
      └─ daily-price-monitor.yml
```

## GitHub Secrets 등록
리포지토리 → Settings → Secrets and variables → Actions → New repository secret

아래 5개를 추가하세요.
- `NAVER_CLIENT_ID`
- `NAVER_CLIENT_SECRET`
- `EMAIL_FROM`
- `EMAIL_TO`
- `EMAIL_APP_PASSWORD`

## Gmail 앱 비밀번호
- Gmail 2단계 인증을 먼저 켜야 합니다.
- 그 다음 Google 계정의 `앱 비밀번호`에서 16자리 비밀번호를 발급해 `EMAIL_APP_PASSWORD`에 넣으세요.

## 실행 방법
### 수동 테스트
- GitHub 저장소 → **Actions**
- **Daily Price Monitor**
- **Run workflow**

### 자동 실행
- 매일 한국시간 오전 9시 실행
- cron: `0 0 * * *` (UTC 00:00 = KST 09:00)

## 결과 확인
- GitHub Actions 실행 상세 화면의 **Artifacts**에서 결과 파일 다운로드
- 이메일 첨부파일로도 수신

## 참고
- GitHub-hosted runner(`ubuntu-latest`)에서는 `win32com`, `os.startfile`, Excel 데스크톱 자동 제어가 동작하지 않으므로 기본값은 꺼져 있습니다.
- VBA 삽입이 꼭 필요하면 `self-hosted` Windows runner로 전환해야 합니다.
