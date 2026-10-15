# 상세 계획 — Google Docs/Slides 기내 편집 도구 (`officesync`)

> 개요는 [PLAN.md](../PLAN.md) 참고. 사용자 1명, 공동 편집 없음, 기내 오프라인 편집이 목적이다.

## 1. 설계

### 1.1 저장소 구조

```
Office/
├─ PLAN.md
├─ docs/
│  ├─ DETAILED_PLAN.md
│  └─ spike-results.md        # 0단계 산출물
├─ pyproject.toml             # uv, ruff, pytest
├─ src/officesync/
│  ├─ cli.py                  # typer 기반 명령
│  ├─ config.py               # 설정 파일
│  ├─ auth.py                 # OAuth, 토큰 저장
│  ├─ drive.py                # URL 해석, 목록, 내보내기, 업로드, 재시도
│  ├─ workspace.py            # 작업 폴더, manifest
│  └─ content.py              # docx/pptx 내용 해시, 편집기 잠금 감지
├─ spikes/                    # 0단계 검증 스크립트
└─ tests/
```

### 1.2 로컬 작업 폴더

```
OfficeSync/                   # 기본: 문서\OfficeSync
├─ 출장보고.docx
├─ 고객 발표.pptx
└─ .officesync/
   ├─ manifest.json
   ├─ backup/<fileId>-<시각>.*   # 덮어쓰기 직전 서버본
   └─ officesync.log
```

### 1.3 `manifest.json`

```json
{
  "schema": 1,
  "files": {
    "1AbCdEf...": {
      "kind": "document",
      "title": "출장보고",
      "local_path": "출장보고.docx",
      "parent_id": "0BxYz...",
      "pulled_at": "2026-10-15T08:30:00+09:00",
      "remote_version": 1532,
      "base_content_hash": "…",
      "last_push": null
    }
  }
}
```

- `remote_version`: Drive `files.version`. push 때 이 값이 그대로면 그사이 서버에서 아무도(다른 기기의 나 포함) 안 고친 것.
- `base_content_hash`: 파일 바이트가 아니라 텍스트+서식 기준 해시. 편집기가 수정 없이 저장만 해도 바이트는 바뀌기 때문.

### 1.4 명령

| 명령 | 설명 | 옵션 |
|---|---|---|
| `login` | 브라우저로 Google 로그인 (최초 1회) | `--client-secret PATH` |
| `pull URL...` | 문서·폴더를 .docx/.pptx로 내려받기 | `--recursive`, `--force` |
| `status` | 수정됨 / 그대로 / 편집기에서 열림 | `--refresh`(서버 확인) |
| `open FILE` | 기본 편집기로 열기 | |
| `push [FILE...]` | 수정한 파일 반영 | `--dry-run`, `--yes` |
| `doctor` | 탑승 전 점검: 토큰, 편집기 설치, 받은 파일 수 | |

종료 코드: 0 성공 / 1 오류 / 3 일부 사본으로 업로드됨 / 4 로그인 필요 / 5 인터넷 없음

### 1.5 push 처리 흐름 (파일 하나)

```
if 편집기에서 열려 있음:          건너뜀 ("닫고 다시 시도")
if 내용 해시 == base:             건너뜀 (수정 없음)

remote = files.get(version, trashed, capabilities.canEdit)

if 삭제됨 / 권한 없음:            내 드라이브에 새 문서로 업로드
elif remote.version == base:      서버본 → backup/ 저장
                                  files.update(로컬 파일 → Google 형식 변환)
                                  manifest 갱신
else:                             "원래이름 (오프라인 편집본 2026-10-15 14시)" 사본 업로드
```

---

## 2. 단계별 작업

시간은 1인 기준 영업일(d). 완료 조건: 코드 + 단위 테스트 + CI 통과 + README 갱신.

### 0단계 — 준비·검증 (2d)

| ID | 작업 | 시간 |
|---|---|---|
| T0.1 | 저장소 골격: `pyproject.toml`, ruff, pytest, GitHub Actions(windows-latest) | 0.5d |
| T0.2 | Google Cloud 프로젝트, Drive API 활성화, "내부" OAuth 데스크톱 클라이언트 생성, 관리자 설정 확인 — [setup-google-cloud.md](setup-google-cloud.md) | 0.5d |
| T0.3 | **V1 내보내기 품질**: 실제로 기내에서 쓸 문서 3~5개를 .docx/.pptx로 받아 편집기에서 확인 (한글 폰트, 표, 이미지, 목록) | 0.5d |
| T0.4 | **V2 덮어쓰기**: 문서 **사본**에 수정한 .docx를 `files.update`로 올려 확인 — 파일 ID·링크 유지, 버전 기록, 서식 | 0.25d |
| T0.5 | **V5 편집기 동작**: 수정 없이 저장 시 변화, 잠금 파일 이름 패턴 | 0.25d |

결과는 `docs/spike-results.md`에 정리.
V1에서 서식 손실이 심하면 → 도구 개발을 멈추고 Google 오프라인 모드 사용을 권장.
V2에서 파일 ID가 바뀌면 → 원본은 두고 새 문서로 올리는 방식으로 변경.

### 1단계 — 내려받기 (4d)

| ID | 작업 | 시간 |
|---|---|---|
| T1.1 | `config.py`: `%APPDATA%\officesync\config.toml` (작업 폴더 위치) | 0.25d |
| T1.2 | `auth.py`: `InstalledAppFlow` 로그인, 토큰은 `keyring`(Windows 자격 증명 관리자), 자동 갱신, 실패 시 종료 코드 4 | 0.75d |
| T1.3 | `drive.py`: URL 해석(문서/프레젠테이션/폴더), 폴더 목록, 바로가기 해석, 내보내기(10MB 넘으면 `exportLinks`), 429/5xx 재시도 | 1d |
| T1.4 | `workspace.py`: manifest 원자적 저장, 파일명 정리(Windows 금지 문자, 같은 이름 → `(2)`) | 0.5d |
| T1.5 | `content.py`: docx/pptx 내용 해시, 잠금 파일 감지 | 0.5d |
| T1.6 | CLI: `login` `pull` `status` `open` `doctor` | 0.5d |
| T1.7 | 단위 테스트 | 0.5d |

완료 기준: 폴더 하나를 받고 인터넷을 끈 채 편집 → `status`가 수정한 파일만 "수정됨"으로 표시.

### 2단계 — 반영 (3d)

| ID | 작업 | 시간 |
|---|---|---|
| T2.1 | `push` 덮어쓰기 경로 (백업 → 업로드 → manifest 갱신) | 1d |
| T2.2 | 사본 경로 (서버도 바뀜 / 삭제됨 / 권한 없음) | 0.5d |
| T2.3 | `--dry-run`, 결과 요약 표, 종료 코드 | 0.5d |
| T2.4 | 시나리오 테스트 (아래 3장) | 1d |

완료 기준: 3장 시나리오 모두 통과 → **v0.1 (기내 사용 가능)**

### 3단계 — 편의 기능 (선택, 2~3d)

| ID | 작업 | 시간 |
|---|---|---|
| T3.1 | PyInstaller로 `officesync.exe` 단일 파일 빌드 (Python 설치 불필요) | 1d |
| T3.2 | 간단한 창: 파일 목록 + [받기] [반영] 버튼 (PySide6 또는 tkinter) | 1~2d |

---

## 3. 테스트 시나리오

| ID | 상황 | 기대 결과 |
|---|---|---|
| S1 | 편집 후 push | 원래 문서에 반영, 링크 그대로 |
| S2 | 기내 편집 중 휴대폰으로 같은 문서 수정 | 원본 유지 + 사본 업로드 |
| S3 | 수정 없이 저장만 | 업로드 안 함 |
| S4 | 편집기에서 연 채로 push | 건너뜀 |
| S5 | 토큰 만료 상태로 push | `login` 안내, 아무것도 안 바뀜 |
| S6 | push 도중 인터넷 끊김 | 원본 손상 없음, 다시 실행하면 완료 |
| S7 | 인터넷 없이 `pull`/`push` | 종료 코드 5, 친절한 안내 |

## 4. 사용 순서 (README에 들어갈 내용)

1. **탑승 전**: `officesync pull <폴더 URL>` → `officesync doctor`로 점검
2. **기내**: `officesync open 출장보고.docx` → 편집 → 저장
3. **착륙 후**: `officesync status` → `officesync push --dry-run` → `officesync push`

## 5. 일정

| 기간 | 내용 |
|---|---|
| 1~2일 차 | 0단계 검증 → 계속할지 결정 |
| 3~6일 차 | 1단계 |
| 7~9일 차 | 2단계 → **v0.1** |
| 10~12일 차 | 3단계 (선택) |
