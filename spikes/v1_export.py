"""V1: 내보내기 품질 확인.

기내에서 편집할 문서를 .docx/.pptx로 받아 구성 요약과 직접 확인할 항목을 report.md로 남긴다.

사용법: uv run python spikes/v1_export.py <문서/프레젠테이션 URL> [<URL> ...]
결과:   spikes/out/v1/
"""

from __future__ import annotations

import argparse
from datetime import datetime

from _common import OUT, credentials, drive_service, export, get_meta
from _inspect import describe

CHECKLIST = """\
## 직접 확인할 것

각 파일을 편집기(ONLYOFFICE·한컴·Word)로 열어 브라우저의 Google 원본과 나란히 비교하세요.

- [ ] 한글 글꼴이 비슷하게 보이는가 (위 "사용 글꼴"이 PC에 없으면 다른 글꼴로 대체됨)
- [ ] 표 모양, 병합된 셀
- [ ] 이미지 위치와 크기
- [ ] 글머리 기호·번호 목록
- [ ] (Docs) 탭이 여러 개인 문서라면 모든 탭의 내용이 들어 있는가
- [ ] (Slides) 도형·텍스트 상자 위치, 줄바꿈, 배경
- [ ] 편집기에서 고치고 저장할 때 오류나 경고가 없는가

판단 기준: 기내에서 편집하기에 불편할 만큼 깨진 문서가 많으면 이 도구 대신
Google 오프라인 모드(Chrome)를 쓰는 편이 낫다.
"""


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("urls", nargs="+", help="Google 문서/프레젠테이션 주소")
    args = parser.parse_args()

    creds = credentials()
    drive = drive_service(creds)
    out = OUT / "v1"

    lines = [f"# V1 내보내기 결과 ({datetime.now():%Y-%m-%d %H:%M})", ""]
    for url in args.urls:
        meta = get_meta(drive, url)
        path = export(drive, creds, meta, out)
        size_kb = path.stat().st_size / 1024
        print(f"받음: {path.name} ({size_kb:,.0f} KB)")

        lines += [
            f"## {meta['name']}",
            "",
            f"- 원본: {meta['webViewLink']}",
            f"- 파일: `{path.name}` ({size_kb:,.0f} KB)",
        ]
        lines += [f"- {key}: {value}" for key, value in describe(path).items()]
        lines.append("")

    lines.append(CHECKLIST)
    report = out / "report.md"
    report.write_text("\n".join(lines), encoding="utf-8")
    print(f"\n보고서: {report}")


if __name__ == "__main__":
    main()
