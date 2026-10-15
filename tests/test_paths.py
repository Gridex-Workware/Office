import pytest

from officesync.paths import safe_filename


@pytest.mark.parametrize(
    ("title", "expected"),
    [
        ("출장보고", "출장보고"),
        ('a/b\\c:d*e?f"g<h>i|j', "a_b_c_d_e_f_g_h_i_j"),
        ("끝에 점.", "끝에 점"),
        ("  앞뒤 공백  ", "앞뒤 공백"),
        ("CON", "_CON"),
        ("con.backup", "_con.backup"),
        ("", "제목 없음"),
        ("///", "___"),
    ],
)
def test_safe_filename(title, expected):
    assert safe_filename(title) == expected


def test_safe_filename_truncates():
    assert len(safe_filename("가" * 500)) == 150
