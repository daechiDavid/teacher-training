# 사용자 직접 검토 체크리스트

## Google Cloud

- OAuth 앱 User type이 `External`인지 확인
- Publishing status가 `Testing`인지 확인
- 실제 사용할 Google 계정이 Test users에 들어 있는지 확인
- OAuth Client type이 `Desktop app`인지 확인
- 등록된 scope가 아래 2개인지 확인
  - `https://www.googleapis.com/auth/drive.file`
  - `https://www.googleapis.com/auth/spreadsheets`

## Google Sheet

- `명단` 탭 이름이 앱 설정값과 정확히 같은지 확인
- 1행 필수 헤더가 모두 있는지 확인
  - `연번`
  - `학교명`
  - `이름`
  - `전화번호`
  - `발급횟수`
  - `과정명1`
  - `발급날짜1`
  - `링크1`
  - `과정명2`
  - `발급날짜2`
  - `링크2`
- 발급횟수와 과정명/발급날짜/링크 기록이 서로 맞는지 확인
- 전화번호가 `010`으로 시작하는 11자리 휴대전화 번호로 정리되어 있는지 확인

## Google Drive

- 영수증 업로드용 상위 폴더를 만들었는지 확인
- 앱 설정에 상위 폴더 ID를 정확히 넣었는지 확인
- Google Workspace 정책상 `링크가 있는 사용자: 뷰어` 공유가 허용되는지 확인

## 앱 설정

- 최종 사용자는 OAuth Client ID/Secret을 입력하지 않음
- 관리자는 개발 실행 전 `src-tauri/google-oauth.example.json`을 참고해 `src-tauri/google-oauth.local.json`을 만들고 실제 Client ID/Secret을 입력
- 배포본은 실제 OAuth 값을 Git에 커밋하지 않고, 앱 데이터 폴더·실행 파일 폴더·배포 리소스 중 하나의 `google-oauth.json` 또는 관리형 설정으로 제공
- 데스크톱 앱에 포함된 Client Secret은 완전한 비밀값이 아니므로, 민감도 요구가 높아지면 별도 토큰 브로커 서버 도입 필요
- 최초 로그인 후 0단계에서 전체명단 Google Sheet URL을 입력하면 앱이 Spreadsheet ID를 추출하는지 확인
- 최초 로그인 후 0단계에서 영수증 업로드 Google Drive 폴더 URL을 입력하면 앱이 폴더 ID를 추출하는지 확인
- 기본 설정 저장 후 다음 앱 실행에서 0단계를 건너뛰고 이수자 명단 업로드 화면으로 진입하는지 확인

## 실제 처리 전

- 이수자 엑셀의 컬럼 순서 확인
  - `연번 | 연수과정명 | 성명 | 학교명 | 연락처 | 접속시간 | 이수결과 | 비고`
- `이수결과`가 `"이수"`인 행만 처리되는지 미리보기에서 확인
- 수동 확인 대상이 0건인지 확인
- 발급 가능 대상자 수가 예상 인원과 맞는지 확인

## 보안

- `.npmrc`의 `ignore-scripts=true` 유지
- `pnpm-lock.yaml`을 기준으로 설치
- 의존성 추가 후 `pnpm audit` 실행
- OAuth Client Secret, refresh token, 발급 PDF를 Git이나 공개 폴더에 올리지 않기
