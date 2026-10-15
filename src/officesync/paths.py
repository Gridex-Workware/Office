"""Google 문서 제목을 Windows에서도 쓸 수 있는 파일 이름으로 바꾼다."""

from __future__ import annotations

import re

_INVALID = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
_MAX_LEN = 150


def safe_filename(title: str) -> str:
    """금지 문자는 `_`로 바꾸고, 끝의 점·공백과 예약어를 피한다. 확장자는 붙이지 않는다."""
    name = _INVALID.sub("_", title).strip().rstrip(". ")
    if not name:
        name = "제목 없음"
    if name.split(".")[0].upper() in _RESERVED:
        name = f"_{name}"
    return name[:_MAX_LEN].rstrip(". ")
