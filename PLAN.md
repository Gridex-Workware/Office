# Google Docs/Slides 오프라인 편집 도구 — 개발 계획

> 작업 단위 상세 계획: [docs/DETAILED_PLAN.md](docs/DETAILED_PLAN.md)

## 1. 목표

인터넷이 없을 때 Google Docs/Slides를 로컬 오피스 프로그램(ONLYOFFICE, 한컴, Word 등)으로 편집하고,
다시 연결되면 원래 Google 문서에 **다른 사람의 수정을 잃지 않고** 반영한다.

**범위 밖**
- 오피스 편집기 자체 개발 (설치된 편집기를 연다)
- 실시간 공동 편집
- Sheets (추후 검토)
- 댓글·제안 모드 동기화

## 2. 전제 (미확정 — 바뀌면 계획 수정)

| 항목 | 기본값 |
|---|---|
| OS | Windows 우선, Mac은 5단계에서 |
| 언어 | Python 3.12 |
| 편집기 | OS 기본 연결 프로그램으로 엶 (ONLYOFFICE 권장, 한컴·Word도 가능) |
| 대상 문서 | 회사 Workspace 비공개 문서 (최초 1회 로그인) + 공개 링크 문서 |
| 공동 편집자 | 있다고 가정 → 덮어쓰기 안전장치 필수 |

## 3. 전체 흐름

```
 온라인                      오프라인                    다시 온라인
 ──────                      ────────                    ───────────
 pull: 내려받기         →    로컬 편집기로 편집     →    push: 반영
  ├ .docx / .pptx 저장       (ONLYOFFICE/한컴/Word)       ├ 서버가 그대로면 → 바로 반영
  └ 원본 스냅샷 보관                                       └ 서버도 바뀌었으면 → 병합, 안 되면 사본 업로드
```

## 4. 단계별 계획

### 0단계 — 준비 및 기술 검증 (1주)

가장 불확실한 부분을 먼저 확인해 이후 설계를 확정한다.

**준비**
- Google Cloud 프로젝트 생성, Drive / Docs / Slides API 활성화
- OAuth 클라이언트(데스크톱 앱) 생성, 동의 화면을 "내부(Internal)"로 설정 → Google 앱 검수 불필요
- Workspace 관리자에게 앱 허용 및 API 접근 정책 확인
- 테스트 문서 세트: 서식·표·이미지·댓글·한글 폰트가 들어간 Docs 3개, Slides 3개

**검증 항목**

| # | 확인할 것 | 결과에 따라 결정되는 것 |
|---|---|---|
| V1 | .docx/.pptx 내보내기 품질 (한글 폰트, 표, 이미지, 도형) | 권장 편집기, 사용자 안내 문구 |
| V2 | 기존 Google 문서에 .docx를 업로드해 덮어쓰면 파일 ID·공유 설정·버전 기록·댓글이 유지되는지 | 2단계 덮어쓰기 방식 (안 되면 새 사본 업로드 방식으로) |
| V3 | Docs API `targetRevisionId`로 예전 버전 기준 수정을 보내면 서버가 병합하는지, 얼마 동안 유효한지 (공식 문서상 수 분) | 3단계 병합 전략 |
| V4 | .pptx로 내보낸 도형을 Slides API `objectId`와 대응시킬 수 있는지 | 4단계 진행 여부 |

**완료 기준**: V1~V4 결과를 `docs/spike-results.md`에 정리, 4단계 진행 여부 결정.

### 1단계 — 내려받기 MVP (1.5주)

- `officesync login` — OAuth 로그인, 토큰은 OS 자격 증명 저장소(Windows 자격 증명 관리자)에 저장
- `officesync pull <문서 URL | 폴더 URL>` — Docs → .docx, Slides → .pptx
- 공개 링크 문서는 로그인 없이 export URL로 받기
- 내려받을 때 병합 기준이 될 원본 스냅샷 보관: 원본 파일 + Docs/Slides API 구조(JSON) + revisionId
- `officesync open <파일>` — 기본 편집기로 열기
- `officesync status` — 로컬에서 수정된 파일 표시 (해시 비교)

**완료 기준**: 폴더 하나(문서 20개)를 받은 뒤 인터넷을 끊고 편집할 수 있다.

### 2단계 — 안전한 업로드 (1주)

`officesync push`
1. 서버 문서가 pull 이후 바뀌었는지 확인 (revisionId / modifiedTime)
2. 안 바뀌었으면 → 서버 현재본을 `backup/`에 내려받은 뒤 덮어쓰기 (V2에서 확인한 방식)
3. 바뀌었으면 → 덮어쓰지 않고 `원래이름 (오프라인 편집본 YYYY-MM-DD)` 사본으로 올리고 알림
4. `--dry-run` — 무엇이 어떻게 올라갈지 미리 보기

**완료 기준**: 어떤 순서로 편집해도 다른 사람의 수정이 사라지지 않는다 (자동 테스트로 증명).

> **여기까지가 실사용 가능한 1차 버전 (약 3.5주)**

### 3단계 — Docs 변경분 병합 (2.5주)

덮어쓰기 대신 바뀐 부분만 반영해 댓글과 다른 사람의 수정을 보존한다.

- 원본 .docx와 편집한 .docx를 문단 단위로 비교 (python-docx)
- 차이를 Docs API 명령으로 변환: `insertText`, `deleteContentRange`, `updateTextStyle`, `updateParagraphStyle`, `createParagraphBullets`
- 서버 반영: push할 때 서버 최신본을 다시 읽고 → 원본·내 편집본·서버 현재본을 로컬에서 3-way 병합
  → 방금 읽은 버전을 `targetRevisionId`로 붙여 전송 (계산하는 동안의 동시 편집은 서버가 병합).
  같은 문단을 양쪽이 고쳤으면 충돌로 보고 2단계 사본 방식으로 대체
  - `targetRevisionId` 병합은 공식 문서상 읽은 뒤 **수 분**만 유효하므로 오프라인 기간 전체를 맡길 수는 없다
- 지원 범위: 텍스트, 굵게/기울임/밑줄, 제목 수준, 목록. 표·이미지 변경은 충돌로 처리
- 참고: [Telofy/DocSync](https://github.com/Telofy/DocSync) (부분 패치, 서버 변경 시 쓰기 거부),
  [kwanUm/obsidian-docs-sync](https://github.com/kwanUm/obsidian-docs-sync) (겹치지 않는 수정 자동 병합, 충돌 표시)

**완료 기준**: 다른 사람이 같은 문서의 다른 문단을 고친 상황에서 양쪽 수정이 모두 남는다.

### 4단계 — Slides 도형 단위 병합 (2.5주, V4 통과 시)

- 원본 .pptx와 편집본 .pptx를 슬라이드·도형 단위로 비교 (python-pptx)
- 도형 ↔ Slides `objectId` 대응 (V4에서 찾은 방법)
- 텍스트 변경 → `deleteText` + `insertText`, `requiredRevisionId`를 붙여 전송
- 거절되면(그사이 서버 변경) → 최신본을 다시 읽고, 내가 고친 도형을 서버에서도 고쳤는지 확인
  → 아니면 새 revisionId로 재시도, 맞으면 충돌
- 지원 범위: 기존 도형의 텍스트 수정, 슬라이드 삭제·순서 변경. 새 도형·이미지 추가는 사본 방식으로

**완료 기준**: 다른 사람이 다른 슬라이드를 고친 상황에서 양쪽 수정이 모두 남는다.

### 5단계 — 앱 형태로 정리 (1.5주)

- 트레이 앱: 문서 목록, 더블클릭으로 열기, 상태 표시 (최신 / 수정됨 / 충돌)
- 인터넷 연결 감지 → 자동 push (사용자 확인 옵션)
- Windows 설치 파일 (PyInstaller), 필요하면 Mac 빌드

## 5. 구조

### 로컬 폴더

```
OfficeSync/
├─ 주간보고.docx
├─ 제안서 발표.pptx
└─ .officesync/
   ├─ manifest.json     # 파일 ID, 이름, 종류, revisionId, 내려받은 시각, 해시
   ├─ base/             # 내려받을 당시 원본 (병합 기준)
   │  ├─ <fileId>.docx
   │  └─ <fileId>.json  # Docs/Slides API 구조 스냅샷
   └─ backup/           # 덮어쓰기 전 서버본 백업
```

### 코드 모듈

```
officesync/
├─ auth.py          # OAuth, 토큰 저장
├─ drive.py         # 목록, 내보내기, 업로드
├─ workspace.py     # 로컬 폴더, manifest, 스냅샷
├─ docs_merge.py    # 3단계
├─ slides_merge.py  # 4단계
├─ cli.py
└─ gui/             # 5단계
```

### 기술 스택

Python 3.12, google-api-python-client, google-auth-oauthlib, python-docx, python-pptx, keyring,
pytest, (5단계) PySide6, PyInstaller

### 권한 (OAuth scope)

- `drive` — 기존 문서 갱신에 필요 (1단계만 쓸 경우 `drive.readonly`로 축소 가능)
- `documents`, `presentations` — 3·4단계

## 6. 위험과 대응

| 위험 | 대응 |
|---|---|
| 덮어쓰기로 다른 사람 수정 소실 | push 전 서버 변경 확인, 변경 시 사본 업로드, 덮어쓰기 전 백업 + Drive 버전 기록 |
| .docx/.pptx 왕복 시 서식 손실 | 0단계에서 손실 목록 정리 후 안내, 3단계부터는 변경분만 반영해 영향 축소 |
| 서버 병합(`targetRevisionId`)은 읽은 뒤 수 분만 유효 | 항상 최신본을 다시 읽어 로컬에서 3-way 병합 후 전송 |
| .pptx 도형과 Slides 객체 대응 불가 | 0단계에서 먼저 검증, 안 되면 Slides는 2단계 방식 유지 |
| 회사 정책으로 API·앱 차단 | 0단계에서 관리자 승인부터 확보 |
| API 할당량 초과 | 변경을 `batchUpdate` 한 번으로 묶고, 실패 시 지수 백오프 재시도 |
| 토큰 유출 | OS 자격 증명 저장소 사용, 평문 파일로 저장하지 않음 |

## 7. 일정 요약 (1인 개발 기준 추정)

| 단계 | 기간 | 누적 |
|---|---|---|
| 0. 준비·검증 | 1주 | 1주 |
| 1. 내려받기 | 1.5주 | 2.5주 |
| 2. 안전한 업로드 | 1주 | 3.5주 ← 1차 사용 가능 |
| 3. Docs 병합 | 2.5주 | 6주 |
| 4. Slides 병합 | 2.5주 | 8.5주 |
| 5. 앱 정리 | 1.5주 | 10주 |

## 8. 결정이 필요한 사항

1. 사용 OS (Windows / Mac / 둘 다)
2. 개발 언어 Python으로 괜찮은지
3. Slides 병합(4단계)까지 필요한지, 2단계 방식(서버가 바뀌면 사본 업로드)으로 충분한지
4. 회사 Workspace 관리자 승인(OAuth 앱 허용)을 받을 수 있는지
