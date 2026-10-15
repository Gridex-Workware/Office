# Google Cloud 설정 (gridex.biz 계정)

officesync가 내 Google Drive 문서를 읽고 쓰려면 Google Cloud 프로젝트와 로그인용 OAuth 클라이언트가 필요합니다.
회사 계정이므로 앱을 **"내부(Internal)"**로 만듭니다. 그러면 Google 앱 검수가 필요 없고, 로그인이 7일마다 풀리지도 않습니다.

소요 시간: 약 15분 (관리자 설정 제외)

## 1. 프로젝트 만들기

1. <https://console.cloud.google.com> 에 **khj@gridex.biz**로 로그인
2. 상단의 프로젝트 선택 → **새 프로젝트**
   - 이름: `officesync`
   - 조직: `gridex.biz` ← 조직이 보이지 않으면 아래 "문제 해결" 참고
3. 만들기

## 2. Drive API 켜기

1. 메뉴 → **API 및 서비스 → 라이브러리** (APIs & Services → Library)
2. `Google Drive API` 검색 → **사용**

Docs API, Slides API는 지금 계획(덮어쓰기 방식)에서는 필요 없습니다.

## 3. OAuth 동의 화면

1. 메뉴 → **Google 인증 플랫폼** (Google Auth Platform) → **시작하기**
2. 입력
   - 앱 이름: `officesync`
   - 사용자 지원 이메일: 본인 주소
   - 대상(Audience): **내부(Internal)**
   - 연락처 이메일: 본인 주소
3. 만들기

## 4. 데스크톱 클라이언트 만들기

1. **Google 인증 플랫폼 → 클라이언트** (Clients) → **클라이언트 만들기**
2. 애플리케이션 유형: **데스크톱 앱** (Desktop app), 이름: `officesync-desktop`
3. 만들기 → 나오는 창에서 **바로 JSON 다운로드**
   (나중에는 클라이언트 보안 비밀번호를 다시 볼 수 없을 수 있습니다)
4. 받은 파일 이름을 `client_secret.json`으로 바꿔 **저장소 최상위 폴더**(README.md가 있는 곳)에 둡니다.
   - `.gitignore`에 등록돼 있어 커밋되지 않습니다. 메신저·메일로 공유하지 마세요.

## 5. Workspace 관리자 설정 확인

관리자 권한이 없다면 이 부분만 관리자에게 요청하세요. <https://admin.google.com>

1. **Drive SDK 허용**
   앱 → Google Workspace → 드라이브 및 문서 → 기능 및 애플리케이션 →
   **Drive SDK** "사용자가 Drive SDK API를 사용하여 Google Drive에 액세스하도록 허용" 켜짐 확인
2. **내부 앱 신뢰** (회사가 서드 파티 앱 접근을 제한한 경우에만)
   보안 → 액세스 및 데이터 제어 → **API 제어** →
   - 설정에서 **"내부 도메인 소유 앱 신뢰"** 체크, 또는
   - 서드 파티 앱 액세스 관리 → 앱 구성 → 4단계의 클라이언트 ID 입력 → **신뢰할 수 있음**

메뉴 이름은 콘솔 버전에 따라 조금 다를 수 있습니다.

## 6. 확인

```powershell
uv run python spikes/v1_export.py <아무 Google 문서 URL>
```

브라우저에 로그인 창이 뜨고, 허용하면 `spikes/out/v1/`에 .docx가 생기면 성공입니다.

## 문제 해결

| 증상 | 원인과 해결 |
|---|---|
| 프로젝트 만들 때 조직에 `gridex.biz`가 없음 | 회사 Google Cloud 조직이 아직 없거나 권한이 없음. 관리자에게 조직 생성/권한 요청. 급하면 "외부(External)·테스트" 앱으로 진행 가능하나 로그인이 7일마다 만료됨 |
| 동의 화면에 "내부"를 고를 수 없음 | 프로젝트가 조직 밖에 있음. 1단계에서 조직을 `gridex.biz`로 지정해 다시 만들기 |
| 로그인 시 "관리자가 차단한 앱" | 5단계 2번 설정 필요 |
| `access_denied` / `accessNotConfigured` | 2단계 Drive API 사용 설정 확인 |
| 다른 계정으로 로그인됨 | `spikes/.token.json` 삭제 후 다시 실행 |
