import pytest

from officesync.urls import DriveRef, parse_drive_url

DOC_ID = "1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789_-ab"


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        (f"https://docs.google.com/document/d/{DOC_ID}/edit", DriveRef("document", DOC_ID)),
        (
            f"https://docs.google.com/document/u/1/d/{DOC_ID}/edit?tab=t.0",
            DriveRef("document", DOC_ID),
        ),
        (
            f"https://docs.google.com/presentation/d/{DOC_ID}/edit#slide=id.p",
            DriveRef("presentation", DOC_ID),
        ),
        (f"https://docs.google.com/spreadsheets/d/{DOC_ID}/edit", DriveRef("spreadsheet", DOC_ID)),
        (f"https://drive.google.com/drive/folders/{DOC_ID}", DriveRef("folder", DOC_ID)),
        (
            f"https://drive.google.com/drive/u/0/folders/{DOC_ID}?usp=sharing",
            DriveRef("folder", DOC_ID),
        ),
        (f"https://drive.google.com/file/d/{DOC_ID}/view", DriveRef("unknown", DOC_ID)),
        (f"https://drive.google.com/open?id={DOC_ID}", DriveRef("unknown", DOC_ID)),
        (f"  {DOC_ID}  ", DriveRef("unknown", DOC_ID)),
    ],
)
def test_parse(text, expected):
    assert parse_drive_url(text) == expected


@pytest.mark.parametrize(
    "text",
    [
        "https://example.com/document/d/1AbCdEfGhIjKlMnOp/edit",
        "https://docs.google.com/document/",
        "짧은id",
    ],
)
def test_parse_rejects(text):
    with pytest.raises(ValueError):
        parse_drive_url(text)
