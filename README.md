# officesync

Google Docs/Slides를 비행기 안(인터넷 없음)에서 ONLYOFFICE·한컴·Word로 편집하고,
착륙 후 원래 Google 문서에 반영하는 도구입니다. 사용자 1명 기준입니다.

- 계획: [PLAN.md](PLAN.md), [docs/DETAILED_PLAN.md](docs/DETAILED_PLAN.md)
- 현재 단계: **0단계 — 준비·검증**

## 0단계 검증 실행 방법 (Windows)

### 1. 준비 (1회)

1. [docs/setup-google-cloud.md](docs/setup-google-cloud.md)를 따라 `client_secret.json`을 만들어
   이 폴더(저장소 최상위)에 둡니다. 이 파일은 `.gitignore`에 들어 있어 커밋되지 않습니다.
2. PowerShell에서 uv와 저장소를 준비합니다.

   ```powershell
   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
   git clone https://github.com/gridex-khj/Office
   cd Office
   git checkout ccr-e0901626-awj6z4
   uv sync
   ```

처음 스크립트를 실행하면 브라우저가 열리고 gridex.biz 계정으로 로그인하게 됩니다.

### 2. 검증 스크립트

| 검증 | 명령 | 하는 일 |
|---|---|---|
| V1 내보내기 품질 | `uv run python spikes/v1_export.py <URL> [<URL> ...]` | 문서를 .docx/.pptx로 받아 `spikes/out/v1/`에 저장하고, 글꼴·표·이미지 수를 `report.md`로 정리 |
| V2 덮어쓰기 | `uv run python spikes/v2_overwrite.py <URL>` | **원본의 사본**을 만들어, 수정한 파일로 덮어쓴 뒤 파일 ID·링크·공유·버전 기록이 유지되는지 확인 |
| V5 편집기 동작 | `uv run python spikes/v5_editor.py <내려받은 파일>` | 편집기로 "수정 없이 저장"과 "한 글자 수정 후 저장"을 해 보고, 수정 여부를 정확히 구분하는지와 잠금 파일 이름을 확인 |

URL에는 기내에서 실제로 편집할 문서·프레젠테이션 주소를 넣으세요.
결과는 [docs/spike-results.md](docs/spike-results.md)에 정리합니다.

## 개발

```bash
uv sync
uv run ruff check . && uv run ruff format --check .
uv run pytest
```
