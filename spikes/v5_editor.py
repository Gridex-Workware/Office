"""V5: 편집기의 저장 동작 확인.

1) 수정 없이 저장했을 때 '수정 없음'으로 판정되는지
2) 한 글자 고쳐 저장했을 때 '수정됨'으로 판정되는지
3) 편집기가 파일을 열고 있는 동안 어떤 잠금 파일을 만드는지

원본 대신 spikes/out/v5/ 에 복사한 파일로 실험한다.

사용법: uv run python spikes/v5_editor.py <내려받은 .docx/.pptx 파일>
"""

from __future__ import annotations

import argparse
import hashlib
import os
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path

from _common import OUT

from officesync.content import fingerprint


def open_in_editor(path: Path) -> None:
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)], check=False)
    else:
        subprocess.run(["xdg-open", str(path)], check=False)


class LockWatcher(threading.Thread):
    """폴더를 0.5초마다 훑어 새로 생겼다 사라지는 파일(잠금 파일)을 기록한다."""

    def __init__(self, folder: Path) -> None:
        super().__init__(daemon=True)
        self.folder = folder
        self.initial = set(os.listdir(folder))
        self.seen: set[str] = set()
        self.stop = threading.Event()

    def run(self) -> None:
        while not self.stop.is_set():
            self.seen |= set(os.listdir(self.folder)) - self.initial
            time.sleep(0.5)


def state(path: Path) -> dict:
    fp = fingerprint(path)
    return {
        "바이트": hashlib.sha256(path.read_bytes()).hexdigest(),
        "텍스트": fp.text,
        "텍스트+서식": fp.format,
    }


def round_trip(path: Path, instruction: str) -> tuple[dict, dict, set[str]]:
    before = state(path)
    watcher = LockWatcher(path.parent)
    watcher.start()
    open_in_editor(path)
    input(f"\n{instruction}\n편집기를 완전히 닫은 뒤 Enter를 누르세요... ")
    watcher.stop.set()
    watcher.join()
    return before, state(path), watcher.seen


def compare(before: dict, after: dict) -> list[str]:
    return [f"  {key}: {'바뀜' if before[key] != after[key] else '그대로'}" for key in before]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("file", type=Path, help="spikes/v1_export.py로 받은 .docx/.pptx")
    args = parser.parse_args()

    work = OUT / "v5"
    work.mkdir(parents=True, exist_ok=True)
    path = work / args.file.name
    shutil.copy2(args.file, path)
    print(f"실험용 사본: {path}")

    b1, a1, locks1 = round_trip(path, "[1/2] 아무것도 고치지 말고 '저장'만 하세요 (Ctrl+S).")
    b2, a2, locks2 = round_trip(path, "[2/2] 본문에서 글자 하나만 고치고 저장하세요.")

    lines = [
        "# V5 편집기 저장 동작",
        "",
        f"파일: {args.file.name}",
        "",
        "## 1) 수정 없이 저장",
        *compare(b1, a1),
        "  기대: 텍스트·텍스트+서식 모두 '그대로' (바이트는 바뀌어도 됨)",
        "",
        "## 2) 한 글자 수정 후 저장",
        *compare(b2, a2),
        "  기대: 텍스트 '바뀜'",
        "",
        "## 편집 중 생긴 파일 (잠금 파일 후보)",
        *(f"  {name}" for name in sorted(locks1 | locks2)),
    ]
    print("\n" + "\n".join(lines))
    report = work / "report.md"
    report.write_text("\n".join(lines), encoding="utf-8")
    print(f"\n보고서: {report}")


if __name__ == "__main__":
    main()
