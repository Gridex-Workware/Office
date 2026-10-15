"""0단계 검증 스크립트 공통 부분: 로그인, Drive 서비스, 메타데이터, 내보내기.

검증용 코드라 토큰은 `spikes/.token.json`에 저장한다 (제품에서는 OS 자격 증명 저장소 사용).
"""

from __future__ import annotations

import sys
from pathlib import Path

from google.auth.exceptions import RefreshError
from google.auth.transport.requests import AuthorizedSession, Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from officesync.paths import safe_filename
from officesync.urls import parse_drive_url

SPIKES = Path(__file__).resolve().parent
CLIENT_SECRET = SPIKES.parent / "client_secret.json"
TOKEN = SPIKES / ".token.json"
OUT = SPIKES / "out"
SCOPES = ["https://www.googleapis.com/auth/drive"]

DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
EXPORTS = {
    "application/vnd.google-apps.document": (DOCX, ".docx"),
    "application/vnd.google-apps.presentation": (PPTX, ".pptx"),
}
SHORTCUT = "application/vnd.google-apps.shortcut"
FIELDS = (
    "id,name,mimeType,version,modifiedTime,webViewLink,exportLinks,parents,"
    "capabilities/canEdit,shortcutDetails"
)


def credentials() -> Credentials:
    creds = None
    if TOKEN.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN), SCOPES)
    if creds and creds.valid:
        return creds
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except RefreshError:
            creds = None
    if not creds or not creds.valid:
        if not CLIENT_SECRET.exists():
            sys.exit(f"{CLIENT_SECRET} 가 없습니다. docs/setup-google-cloud.md 를 참고하세요.")
        flow = InstalledAppFlow.from_client_secrets_file(str(CLIENT_SECRET), SCOPES)
        creds = flow.run_local_server(port=0)
    TOKEN.write_text(creds.to_json(), encoding="utf-8")
    return creds


def drive_service(creds: Credentials):
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def get_meta(drive, url_or_id: str) -> dict:
    """URL을 해석해 메타데이터를 가져온다. 바로가기면 원본을 따라간다."""
    ref = parse_drive_url(url_or_id)
    meta = drive.files().get(fileId=ref.id, fields=FIELDS, supportsAllDrives=True).execute()
    if meta["mimeType"] == SHORTCUT:
        target = meta["shortcutDetails"]["targetId"]
        meta = drive.files().get(fileId=target, fields=FIELDS, supportsAllDrives=True).execute()
    if meta["mimeType"] not in EXPORTS:
        sys.exit(f"Google 문서/프레젠테이션이 아닙니다: {meta['name']} ({meta['mimeType']})")
    return meta


def export(drive, creds: Credentials, meta: dict, dest_dir: Path, suffix: str = "") -> Path:
    """Google 문서를 .docx/.pptx로 저장한다. 10MB를 넘으면 exportLinks로 받는다."""
    mime, ext = EXPORTS[meta["mimeType"]]
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / f"{safe_filename(meta['name'])}{suffix}{ext}"
    try:
        data = drive.files().export(fileId=meta["id"], mimeType=mime).execute()
    except HttpError as e:
        if b"exportSizeLimitExceeded" not in (e.content or b""):
            raise
        response = AuthorizedSession(creds).get(meta["exportLinks"][mime])
        response.raise_for_status()
        data = response.content
    dest.write_bytes(data)
    return dest
