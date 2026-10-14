# 상세 계획 — Google Docs/Slides 오프라인 편집 도구 (`officesync`)

> 개요는 [PLAN.md](../PLAN.md) 참고. 이 문서는 작업 단위(태스크), 데이터 구조, 알고리즘, 테스트, 일정을 다룬다.

## 0. 전제와 이전 계획에서 바뀐 점

**전제 (미확정, 바뀌면 이 문서 수정)**
- Windows 우선, Python 3.12, CLI 먼저 → 5단계에서 트레이 앱
- 편집은 OS 기본 연결 프로그램 (ONLYOFFICE 권장, 한컴·Word 지원 목표)
- 회사 Workspace 비공개 문서 + 공개 링크 문서, 공동 편집자 있음

**바뀐 점: Docs 서버 병합 기능의 유효 시간**
- Docs API `writeControl.targetRevisionId`(서버가 다른 사람 수정과 병합해 주는 기능)는
  공식 문서상 **읽은 뒤 수 분 동안만** 쓸 수 있다. (24시간은 `revisionId` 값 자체의 보장 기간)
- 따라서 오프라인 편집(보통 수 시간~수 일)에는 직접 쓸 수 없다.
- **3단계 설계**: push할 때 항상 서버 최신본을 다시 읽고 → **로컬에서 3-way 병합** →
  방금 읽은 버전을 `targetRevisionId`로 붙여 전송 (계산하는 수 초 사이의 동시 편집만 서버가 병합).
- Slides API는 `requiredRevisionId`만 있다 (버전이 다르면 거절). 거절되면 다시 읽고 재계산.

---

## 1. 공통 설계

### 1.1 저장소 구조

```
Office/
├─ PLAN.md
├─ docs/
│  ├─ DETAILED_PLAN.md
│  └─ spike-results.md          # 0단계 산출물
├─ pyproject.toml               # uv, ruff, pytest 설정
├─ src/officesync/
│  ├─ cli.py                    # typer 기반 명령
│  ├─ config.py                 # 설정 파일
│  ├─ auth.py                   # OAuth, 토큰 저장
│  ├─ gapi.py                   # API 클라이언트 생성, 공통 재시도
│  ├─ drive.py                  # URL 해석, 목록, 내보내기, 업로드
│  ├─ workspace.py              # 작업 폴더, manifest, 스냅샷
│  ├─ status.py                 # 파일 상태 판정, 편집기 잠금 감지
│  ├─ push.py                   # 업로드 전략 선택 (덮어쓰기 / 병합 / 사본)
│  ├─ model/                    # 문서 공통 중간 모델
│  │  ├─ doc_model.py
│  │  ├─ from_docx.py
│  │  ├─ from_gdoc.py
│  │  ├─ slide_model.py
│  │  ├─ from_pptx.py
│  │  └─ from_gslides.py
│  ├─ merge/
│  │  ├─ align.py               # 같은 버전의 두 표현 정렬
│  │  ├─ diff3.py               # 3-way 비교
│  │  ├─ docs_ops.py            # 변경 → Docs API 요청
│  │  └─ slides_ops.py          # 변경 → Slides API 요청
│  └─ gui/                      # 5단계
├─ spikes/                      # 0단계 검증 스크립트 (제품 코드 아님)
└─ tests/
   ├─ unit/
   ├─ fixtures/                 # docx/pptx/API JSON 샘플
   └─ integration/              # 실제 Google 계정 필요, 수동 실행
```

### 1.2 로컬 작업 폴더

```
OfficeSync/                       # 사용자가 지정 (기본: 문서\OfficeSync)
├─ 영업팀/주간보고.docx           # Drive 폴더 구조 유지
├─ 제안서 발표.pptx
└─ .officesync/
   ├─ manifest.json
   ├─ base/<fileId>.docx|.pptx    # 내려받을 당시 원본 (병합 기준 O)
   ├─ base/<fileId>.json          # 같은 시점의 Docs/Slides API 응답
   ├─ backup/<fileId>-<시각>.*    # 덮어쓰기 직전 서버본
   ├─ conflicts/<제목>-<시각>.md  # 충돌 보고서
   └─ log/officesync.log
```

### 1.3 `manifest.json`

```json
{
  "schema": 1,
  "account": "user@example.com",
  "files": {
    "1AbCdEf...": {
      "kind": "document",
      "title": "주간보고",
      "local_path": "영업팀/주간보고.docx",
      "parent_id": "0BxYz...",
      "pulled_at": "2026-10-14T09:12:00+09:00",
      "remote_version": 1532,
      "remote_modified_time": "2026-10-14T00:10:22.000Z",
      "revision_id": "ALm37BW...",
      "base_sha256": "…",
      "base_content_hash": "…",
      "can_edit": true,
      "last_push": { "at": "…", "result": "overwrite|merged|copy", "copy_id": null }
    }
  }
}
```

- `remote_version`: Drive `files.version` (서버 변경 시 증가). 변경 감지의 1차 기준.
- `base_content_hash`: 파일 바이트가 아니라 **내용**(텍스트+서식) 해시.
  편집기가 수정 없이 "저장"만 해도 zip 바이트가 바뀌므로, 수정 여부는 이 값으로 판단.
- 저장은 항상 임시 파일에 쓴 뒤 이름 바꾸기(원자적 저장).

### 1.4 CLI 명세

| 명령 | 설명 | 주요 옵션 |
|---|---|---|
| `login` | 브라우저로 Google 로그인, 토큰 저장 | `--client-secret PATH` |
| `logout` | 저장된 토큰 삭제 | |
| `init [DIR]` | 작업 폴더 생성 | |
| `pull URL...` | 문서·폴더 내려받기 | `--recursive`, `--public`(로그인 없이), `--force`(로컬 수정 덮어쓰기) |
| `status` | 파일별 상태 표 | `--refresh`(서버 확인), `--json` |
| `open FILE` | 기본 편집기로 열기 | |
| `push [FILE...]` | 수정 사항 반영 | `--dry-run`, `--mode auto\|overwrite\|merge\|copy`, `--partial`, `--yes` |
| `resolve FILE` | 충돌 처리 완료 표시 | `--keep local\|remote` |
| `doctor` | 환경 점검 (토큰, API, 편집기, 잠금 파일) | |

**종료 코드**: 0 성공 / 1 오류 / 2 잘못된 사용 / 3 일부 충돌(사본 업로드됨) / 4 로그인 필요 / 5 인터넷 없음

### 1.5 파일 상태 모델

| 로컬 \ 서버 | 그대로 | 바뀜 | 삭제됨 / 권한 없음 |
|---|---|---|---|
| **그대로** | ✅ 최신 | ⬇ 서버본 받기 가능 | ⚠ 경고만 |
| **수정됨** | ⬆ push 가능 (덮어쓰기) | ⚠ 병합 필요 (3단계 전: 사본) | ⚠ 새 문서로 올릴지 확인 |
| **로컬 파일 없음** | manifest에서 뺄지 확인 | ⬇ 다시 받기 가능 | 정리 |

추가 상태: 🔒 **편집기에서 열려 있음** (잠금 파일 감지) → push 건너뜀.

---

## 2. 단계별 작업

표의 시간은 1인 기준 영업일(d). 각 태스크의 완료 조건은 공통으로
"코드 + 단위 테스트 + CI 통과 + README 사용법 갱신"이다.

### 0단계 — 준비 및 기술 검증 (5d)

| ID | 작업 | 산출물 | 시간 |
|---|---|---|---|
| T0.1 | 저장소 골격: `pyproject.toml`(uv), ruff, pytest, GitHub Actions(windows-latest, ubuntu-latest) | CI가 도는 빈 패키지 | 0.5d |
| T0.2 | Google Cloud 프로젝트, Drive·Docs·Slides API 활성화, OAuth 동의 화면 "내부", 데스크톱 클라이언트 생성 | `client_secret.json` (저장소에 넣지 않음) | 0.5d |
| T0.3 | Workspace 관리자 확인: API 접근 제어에서 앱 허용 | 승인 | 대기 |
| T0.4 | 테스트 문서 세트 작성 (아래 표) | 공유 폴더 "officesync-test" | 0.5d |
| T0.5 | **V1** 내보내기 품질 | 체크리스트 결과 | 0.5d |
| T0.6 | **V2** 기존 문서 덮어쓰기 | 결과 + 결정 | 0.5d |
| T0.7 | **V3** 동시성·변경 감지 | 결과 + 결정 | 0.5d |
| T0.8 | **V4** Slides 도형 ID 대응 | 결과 + 4단계 진행 여부 | 0.5d |
| T0.9 | **V5** 편집기 저장 동작 | 결과 | 0.5d |
| T0.10 | `docs/spike-results.md` 정리, 이 계획 수정 | 문서 | 1d |

**테스트 문서 세트 (T0.4)**

| 이름 | 내용 |
|---|---|
| D1 | 제목·본문·굵게/기울임·링크·글머리 목록·번호 목록, 한글+영문+이모지 |
| D2 | 표(병합 셀 포함), 이미지, 각주, 머리글/바닥글 |
| D3 | 탭 여러 개, 긴 문서(50쪽 이상) |
| S1 | 텍스트 상자 위주 10장, 발표자 노트 |
| S2 | 표, 이미지, 도형 그룹, 차트 |
| S3 | 마스터·레이아웃·테마 사용, 한글 폰트 |

**검증 절차**

- **V1 내보내기 품질** — `spikes/v1_export.py`로 6개 문서를 .docx/.pptx로 내보낸 뒤,
  ONLYOFFICE·한컴·Word에서 열어 확인: 폰트 대체, 표, 이미지, 목록 번호, 탭 처리(모든 탭이 들어가는지), 파일 크기.
- **V2 덮어쓰기** — `spikes/v2_overwrite.py`: 테스트 문서 **사본**에 `files.update`로 .docx 업로드.
  확인: 파일 ID 유지 / 공유 설정 유지 / 버전 기록에 새 버전 생김 / 댓글 유지 여부 / 제목 유지 / 서식.
  - 실패 시 대안: 새 문서로 올리고 원본에 안내 문구 추가, 또는 Docs API로 본문 전체 교체.
- **V3 동시성** — `spikes/v3_revision.py`:
  (a) 오래된 `requiredRevisionId` → 400 확인
  (b) 읽은 뒤 다른 계정이 편집 → 1분/5분/30분 후 `targetRevisionId`로 쓰기 → 병합 여부 기록
  (c) 댓글만 달거나 제목만 바꿔도 Drive `version`이 오르는지 (오르면 불필요한 "서버 변경" 판정 → Docs `revisionId` 비교로 보완)
- **V4 Slides 도형 ID** — `spikes/v4_slides_ids.py`: .pptx로 내보낸 뒤 python-pptx로 도형 이름·ID를 덤프하고
  `presentations.get`의 `objectId`와 비교. 도형 이름에 원본 objectId가 남는지,
  편집기로 저장한 뒤에도 유지되는지 확인.
  - 실패 시 대안: 슬라이드 순서 + 도형 순서 + 위치·크기 + 텍스트 유사도로 추정 대응 (정확도 측정 후 판단)
- **V5 편집기 저장 동작** — 수정 없이 저장했을 때 바이트·내용 해시 변화, 잠금 파일 이름 패턴
  (Word `~$이름.docx`, LibreOffice `.~lock.이름.docx#` 등), 스타일 이름 변화(예: 한국어 "제목 1").

**진행/중단 기준**
- V2 실패 → 2단계 덮어쓰기 방식을 대안으로 교체
- V4 실패 + 추정 대응 정확도 95% 미만 → 4단계 취소, Slides는 2단계 방식 유지
- 관리자 승인 불가 → 개인 계정으로 개발하고 승인은 병행 추진

### 1단계 — 내려받기 (7d)

| ID | 작업 | 상세 | 시간 |
|---|---|---|---|
| T1.1 | `config.py` | `%APPDATA%\officesync\config.toml`: 작업 폴더, 편집기 경로(선택), 로그 수준 | 0.5d |
| T1.2 | `auth.py` | `InstalledAppFlow.run_local_server`로 로그인. 토큰은 `keyring`(Windows 자격 증명 관리자, 항목당 약 2.5KB 제한 고려). 만료 시 자동 갱신, 실패 시 종료 코드 4 | 1d |
| T1.3 | `gapi.py` | 서비스 객체 생성, 429/5xx 지수 백오프+지터(최대 5회), 인터넷 없음 감지 → 종료 코드 5 | 0.5d |
| T1.4 | `drive.py` 조회 | URL 해석(문서/프레젠테이션/폴더/ID), `files.get` 필드 지정, 폴더 목록(`files.list`, 공유 드라이브 `supportsAllDrives`), 바로가기 대상 해석 | 1d |
| T1.5 | `drive.py` 내보내기 | `files.export`(10MB 제한) → 넘으면 `exportLinks` URL로 받기. 공개 문서는 export URL로 로그인 없이 | 0.5d |
| T1.6 | `workspace.py` | `init`, manifest 원자적 저장, 파일명 정리(Windows 금지 문자 `\/:*?"<>\|`, 같은 이름 → `(2)`), 폴더 구조 유지, base 스냅샷(파일 + API JSON) 저장 | 1d |
| T1.7 | 내용 해시 | `model/from_docx.py`, `from_pptx.py` 최소 버전(텍스트·서식 → 정규화 → 해시) | 1d |
| T1.8 | CLI | `login` `logout` `init` `pull` `status` `open` `doctor` | 1d |
| T1.9 | 테스트 | URL 해석, 파일명, manifest, 내용 해시 단위 테스트 + 통합 테스트 1개 | 0.5d |

**완료 기준**: 폴더 하나(문서 20개)를 받고, 인터넷을 끊은 상태에서 편집한 뒤 `status`에 "수정됨"으로 정확히 표시된다.
수정 없이 저장만 한 파일은 "그대로"로 표시된다.

### 2단계 — 안전한 업로드 (5d)

**push 처리 흐름 (파일 하나 기준)**

```
if 편집기에서 열려 있음:            건너뜀 ("닫고 다시 시도")
if 로컬 내용 해시 == base:          건너뜀 (수정 없음)

remote = files.get(version, trashed, capabilities.canEdit, name)

if 삭제됨 또는 찾을 수 없음:        사용자 확인 → 새 문서로 업로드
elif 편집 권한 없음:                내 드라이브에 사본 업로드
elif remote.version == base.version:
    서버 현재본 내보내기 → backup/
    files.update(fileId, 로컬 파일, Google 형식으로 변환)    # V2 결과 방식
    다시 내려받아 base 스냅샷 갱신
    결과: overwrite
else:  # 서버도 바뀜
    mode == merge (3단계 이후) → 병합 시도, 실패 시 사본
    그 외 → "원래이름 (오프라인 편집본 2026-10-14 15시30분)" 사본을 같은 폴더에 업로드
            conflicts/ 에 기록, 링크 출력
    결과: copy
```

남는 위험: version 확인과 업로드 사이 수 초 동안 다른 사람이 편집하면 그 수정이 덮일 수 있다.
→ 직전 백업 + Drive 버전 기록으로 복구할 수 있고, 3단계 병합 방식에서는 해소된다.

| ID | 작업 | 시간 |
|---|---|---|
| T2.1 | `status --refresh` (서버 version 비교) | 0.5d |
| T2.2 | 편집기 잠금 감지 (V5 패턴 + Windows 배타적 열기 시도) | 0.5d |
| T2.3 | 덮어쓰기 경로: 백업 → 업로드 → base 갱신 | 1d |
| T2.4 | 사본 경로 + `conflicts/` 기록 | 0.5d |
| T2.5 | 삭제·권한 없음·서버 측 이름 변경 처리 | 0.5d |
| T2.6 | `--dry-run`, 결과 요약 표, 종료 코드 | 0.5d |
| T2.7 | `resolve` 명령 | 0.5d |
| T2.8 | 시나리오 통합 테스트 S1~S8 (4장) | 1d |

**완료 기준**: S1~S8에서 다른 사람의 수정이 한 번도 사라지지 않는다. → **v0.1 배포 (1차 사용 가능)**

### 3단계 — Docs 변경분 병합 (13d)

**핵심 아이디어: 같은 표현끼리만 비교한다**

```
            내려받을 때(O)                    지금
 docx 표현   O_docx  ──── 내 변경 ────▶  A_docx   (내가 편집한 파일)
               │ align (같은 버전이라 거의 1:1)
 API 표현    O_gdoc  ──── 남의 변경 ──▶  B_gdoc   (push 직전 다시 읽은 서버본)
```

- **내 변경** = diff(O_docx, A_docx): 둘 다 docx라 내보내기 변환 차이로 인한 가짜 변경이 없다.
- **남의 변경** = diff(O_gdoc, B_gdoc): 둘 다 API JSON.
- **연결고리** = align(O_docx, O_gdoc): 같은 시점의 두 표현을 문단 단위로 짝지음.
  이를 통해 "내가 고친 O 문단"이 B의 어느 위치(UTF-16 인덱스)인지 찾는다.

**중간 모델 (`doc_model.py`)**

```
Paragraph:
  text: str
  runs: [Run(start, end, bold, italic, underline, strike, link)]
  named_style: NORMAL_TEXT | TITLE | SUBTITLE | HEADING_1..6
  list: None | (kind: bullet|number, level)
  container: body | table(t, row, col)
  # from_gdoc 전용: start_index, end_index (UTF-16)
Opaque: 이미지, 표 구조(행×열), 각주, 수식, 쪽 나눔, 구역 나눔
```

- `from_docx`: 제목 판정은 스타일 이름 + 개요 수준(outlineLvl)을 함께 본다 (한컴·한국어 Word 대비).
- 정규화: NBSP→공백, 줄 안 나눔(`\v`)·탭·스마트 따옴표 통일.

**병합 규칙 (문단 단위 diff3)**

| O→A (나) | O→B (남) | 처리 |
|---|---|---|
| 변경 | 그대로 | 적용 |
| 그대로 | 변경 | 아무것도 안 함 (서버에 이미 있음) |
| 변경 | 변경, 결과 같음 | 아무것도 안 함 |
| 변경 | 변경, 결과 다름 | **충돌** |
| 표 구조·이미지 등 Opaque 변경 | — | **지원 안 함 → 충돌** |

- 충돌이 하나라도 있으면 기본은 **문서 전체 미적용 + 사본 업로드 + 충돌 보고서**.
  `--partial`이면 충돌 없는 부분만 적용하고 보고서를 남긴다.
- 충돌 보고서(`conflicts/*.md`): 문단별로 "원본 / 내 수정 / 서버 수정" 3칸.

**요청 생성 (`docs_ops.py`)**

- 문단 안의 변경은 글자 단위 diff → `deleteContentRange` + `insertText`
- 서식: `updateTextStyle`(fields 마스크 지정), 제목 수준: `updateParagraphStyle(namedStyleType)`,
  목록: `createParagraphBullets` / `deleteParagraphBullets`, 링크: `updateTextStyle(link)`
- 새 문단은 `"\n"`을 포함한 `insertText`. 본문 마지막 줄바꿈은 지우지 않는다.
- **문서 끝에서 앞쪽 순서로** 요청을 쌓아 앞 요청이 뒤 요청의 인덱스를 밀지 않게 한다.
- 인덱스는 UTF-16 단위: `len(s.encode("utf-16-le")) // 2` (이모지·일부 한자 주의).
- 탭이 여러 개인 문서: 요청에 `tabId` 지정 (V1에서 내보내기가 탭을 어떻게 다루는지에 따라 범위 결정).

**전송과 검증**

1. `documents.get` (B, `includeTabsContent=true`) → 병합 계산 → `batchUpdate`
   (`writeControl.targetRevisionId = B.revisionId`, 요청 전체가 원자적으로 적용됨)
2. 400 오류(버전 만료 등) → B를 다시 읽고 재계산, 최대 3회
3. 적용 후 다시 읽어 기대 결과와 비교 → 다르면 경고 + Drive 버전 기록 안내
4. 성공 시 base 스냅샷을 새 버전으로 갱신

| ID | 작업 | 시간 |
|---|---|---|
| T3.1 | `doc_model.py`, `from_gdoc.py` (인덱스 포함) | 1.5d |
| T3.2 | `from_docx.py` 완성 (서식·목록·표 셀·Opaque) | 1.5d |
| T3.3 | `align.py` + 정렬 실패율 측정 (실패율 기준 넘으면 병합 포기 → 사본) | 1.5d |
| T3.4 | `diff3.py` (문단 단위) | 1.5d |
| T3.5 | `docs_ops.py` 텍스트 변경 → 요청 | 2d |
| T3.6 | `docs_ops.py` 서식·제목·목록·링크 | 1.5d |
| T3.7 | 전송·재시도·검증, `push --mode merge` 연결 | 1d |
| T3.8 | 충돌 보고서, `--partial` | 0.5d |
| T3.9 | 테스트: 요청 생성 golden 테스트, 통합 S2·S3·S9·S10 | 2d |

**완료 기준**: 다른 사람이 같은 문서의 다른 문단을 고친 상황(S2)에서 양쪽 수정이 모두 남고 댓글이 유지된다.
→ **v0.2 배포**

### 4단계 — Slides 도형 단위 병합 (12d, V4 통과 시)

**중간 모델 (`slide_model.py`)**

```
Slide:  object_id, order, shapes: [Shape], notes: [Paragraph]
Shape:  object_id, kind(text_box|shape|table|image|group|chart), bbox,
        paragraphs: [Paragraph]   # 표는 셀별
```

**지원 범위 (v1)**

| 변경 | API 요청 | 지원 |
|---|---|---|
| 기존 도형·표 셀 텍스트 수정 | `deleteText(ALL)` + `insertText` + `updateTextStyle` | ✅ |
| 발표자 노트 수정 | 노트 도형에 동일 | ✅ |
| 슬라이드 삭제 | `deleteObject` | ✅ |
| 슬라이드 순서 변경 | `updateSlidesPosition` | ✅ |
| 도형 이동·크기 변경 | `updatePageElementTransform` | 다음 버전 |
| 새 슬라이드·도형·이미지 추가 | — | ❌ 충돌 → 사본 |

**병합 규칙**: 도형 단위 diff3. 같은 도형을 양쪽이 고쳤으면 충돌.
**전송**: `requiredRevisionId = B.revisionId`. 400이면 다시 읽고, 내가 건드린 도형을 서버에서도 건드렸는지 확인 →
아니면 재계산·재시도(최대 3회), 맞으면 충돌.

| ID | 작업 | 시간 |
|---|---|---|
| T4.1 | `slide_model.py`, `from_gslides.py` | 1.5d |
| T4.2 | `from_pptx.py` + 도형 ID 대응 (V4 방법) | 2d |
| T4.3 | 도형 단위 diff3 | 1.5d |
| T4.4 | `slides_ops.py` 텍스트·노트 | 2d |
| T4.5 | `slides_ops.py` 슬라이드 삭제·순서 | 1d |
| T4.6 | 전송·재시도·검증 | 1d |
| T4.7 | 테스트 (golden + 통합 S2·S3 Slides판) | 2d |
| T4.8 | 편집기 3종으로 저장한 pptx 호환 확인 | 1d |

**완료 기준**: 다른 사람이 다른 슬라이드를 고친 상황에서 양쪽 수정이 모두 남는다. → **v0.3 배포**

### 5단계 — 앱 정리 (8d)

| ID | 작업 | 시간 |
|---|---|---|
| T5.1 | PySide6 트레이 앱 골격, CLI 기능 재사용 | 1d |
| T5.2 | 문서 목록 창: 상태 아이콘, 더블클릭 열기, 우클릭(받기/반영/충돌 보기) | 2d |
| T5.3 | 인터넷 연결 감지 → 알림, 자동 push 옵션(기본 꺼짐) | 1d |
| T5.4 | PyInstaller 빌드 + 설치 마법사(Inno Setup), 시작 프로그램 등록 | 1.5d |
| T5.5 | 한국어 사용 설명서 (설치, 로그인, 충돌 처리) | 1d |
| T5.6 | 사내 시범 사용 + 피드백 반영 | 1.5d |

→ **v1.0 배포**

---

## 3. 예외 상황 처리

| 상황 | 처리 |
|---|---|
| 서버에서 문서 삭제(휴지통) | push 시 "새 문서로 올릴까요?" 확인 |
| 편집 권한이 보기 전용으로 바뀜 | 내 드라이브에 사본 업로드 |
| 서버에서 제목 변경 | manifest 제목 갱신, 로컬 파일명은 유지 (`status`에 표시) |
| 로컬에서 파일명 변경 | 같은 내용 해시의 새 파일을 찾아 연결 제안, 못 찾으면 "로컬 없음" |
| 로컬에서 새 .docx/.pptx 생성 | `push --new`로 같은 폴더에 Google 문서로 업로드 (선택 기능) |
| 편집기에서 파일이 열려 있음 | 건너뜀, "닫고 다시 시도" 안내 |
| push 도중 인터넷 끊김 | 문서 단위로 원자적(batchUpdate 1회 / 업로드 1회) → 재실행하면 이어서 처리 |
| 10MB 넘는 문서 | `exportLinks`로 내보내기 |
| 같은 이름 문서 여러 개 | `이름 (2).docx` 식으로 저장, manifest는 파일 ID로 관리 |
| 공유 드라이브 문서 | `supportsAllDrives=true` 지정 |
| 바로가기(shortcut) | 원본 파일 ID로 해석 |
| 토큰 만료·취소 | 자동 갱신, 실패 시 `login` 안내 (종료 코드 4) |
| API 할당량 초과(429) | 지수 백오프 재시도, 계속되면 남은 파일 목록 출력 |

---

## 4. 테스트 계획

**단위 테스트 (CI, 매 push)**
- 모델 변환: python-docx로 만든 샘플 + 실제 내보낸 .docx/.pptx 샘플(`tests/fixtures`)
- align, diff3, 요청 생성: 입력 → 기대 요청 JSON (golden 파일)
- UTF-16 인덱스: 이모지·결합 문자 포함 케이스
- 목표: `merge/` 커버리지 80% 이상
- 실행 환경: windows-latest, ubuntu-latest

**통합 테스트 (실제 Google 계정, 수동 실행)**
- 계정 2개 사용: A(도구 사용자), B("다른 사람" 역할, API로 자동 편집)

| ID | 시나리오 | 기대 결과 |
|---|---|---|
| S1 | A만 수정 | 덮어쓰기(2단계) / 병합(3단계) 성공 |
| S2 | B가 다른 문단 수정 | 2단계: 사본 업로드, 3단계: 양쪽 반영 |
| S3 | A·B가 같은 문단 수정 | 충돌 보고서 + 사본, 원본 변화 없음 |
| S4 | 서버 문서 삭제 | 확인 후 새 문서 |
| S5 | A 권한을 보기 전용으로 변경 | 내 드라이브에 사본 |
| S6 | 편집기에서 열린 채 push | 건너뜀 |
| S7 | 수정 없이 저장만 | "수정 없음" |
| S8 | push 도중 네트워크 차단 | 원본 손상 없음, 재실행 시 완료 |
| S9 | 10MB 넘는 문서 | 내보내기·병합 성공 |
| S10 | 이모지·한자 문단 수정 | 인덱스 어긋남 없음 |

**편집기 호환 매트릭스 (수동)**: ONLYOFFICE / 한컴 / Word × D1~D3, S1~S3 에서 "열기 → 수정 → 저장 → push" 확인.

---

## 5. 일정

| 주차 | 단계 | 주요 결과 |
|---|---|---|
| 1주 | 0단계 | 검증 결과, 계획 확정, 관리자 승인 요청 |
| 2~3주 | 1단계 | `pull` `status` `open` 동작 |
| 3~4주 | 2단계 | **v0.1** — 안전한 덮어쓰기/사본 업로드 (1차 사용 가능) |
| 5~6주 | 3단계 | **v0.2** — Docs 변경분 병합 |
| 7~9주 | 4단계 | **v0.3** — Slides 병합 (V4 통과 시) |
| 9~10주 | 5단계 | **v1.0** — 트레이 앱, 설치 파일 |

합계 약 50 영업일(10주). 4단계를 취소하면 약 7.5주.

**배포 전 체크리스트 (각 버전)**
- [ ] 단위 테스트·CI 통과
- [ ] 해당 단계 통합 시나리오 통과
- [ ] 편집기 호환 매트릭스 확인
- [ ] README·설명서 갱신, 변경 사항 기록
- [ ] 테스트 문서가 아닌 실제 문서 사본으로 1회 이상 시범 사용

---

## 6. 결정이 필요한 사항

| 항목 | 현재 기본값 | 영향 |
|---|---|---|
| 사용 OS | Windows | Mac 포함 시 5단계 +3d |
| 개발 언어 | Python | 변경 시 전체 재추정 |
| Slides 병합 필요 여부 | 필요 (V4 통과 시) | 불필요하면 2.5주 단축 |
| Workspace 관리자 승인 | 받을 수 있다고 가정 | 불가 시 회사 문서 대상 사용 불가 |
| 충돌 시 기본 동작 | 문서 전체 미적용 + 사본 | `--partial`을 기본으로 할지 |
