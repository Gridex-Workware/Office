"""Google 문서·프레젠테이션·폴더 URL에서 종류와 ID를 꺼낸다."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Literal
from urllib.parse import parse_qs, urlparse

Kind = Literal["document", "presentation", "spreadsheet", "folder", "unknown"]

_ID = r"[A-Za-z0-9_-]{10,}"
_PATTERNS: list[tuple[Kind, re.Pattern[str]]] = [
    ("document", re.compile(rf"/document/(?:u/\d+/)?d/({_ID})")),
    ("presentation", re.compile(rf"/presentation/(?:u/\d+/)?d/({_ID})")),
    ("spreadsheet", re.compile(rf"/spreadsheets/(?:u/\d+/)?d/({_ID})")),
    ("folder", re.compile(rf"/drive/(?:u/\d+/)?folders/({_ID})")),
    ("unknown", re.compile(rf"/file/(?:u/\d+/)?d/({_ID})")),
]
_BARE_ID = re.compile(rf"^{_ID}$")


@dataclass(frozen=True)
class DriveRef:
    kind: Kind
    id: str


def parse_drive_url(text: str) -> DriveRef:
    """URL 또는 ID 문자열을 해석한다. 해석할 수 없으면 ValueError."""
    text = text.strip()
    if _BARE_ID.match(text):
        return DriveRef("unknown", text)

    url = urlparse(text)
    if url.hostname not in ("docs.google.com", "drive.google.com"):
        raise ValueError(f"Google 문서 주소가 아닙니다: {text}")

    for kind, pattern in _PATTERNS:
        if m := pattern.search(url.path):
            return DriveRef(kind, m.group(1))

    ids = parse_qs(url.query).get("id")
    if ids and _BARE_ID.match(ids[0]):
        return DriveRef("unknown", ids[0])

    raise ValueError(f"주소에서 문서 ID를 찾지 못했습니다: {text}")
