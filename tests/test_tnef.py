from pathlib import Path

import pytest

from tnefparse.tnef import TNEF

HERE = Path(__file__).resolve().parent


def test_tnef_str_representation():
    tnef_data = (HERE / "examples" / "two-files.tnef").read_bytes()
    t = TNEF(tnef_data)

    assert str(t) == "<TNEF:0x237, 2 attachments>"


def test_tnef_msg_embed():
    tnef_data = (HERE / "examples" / "IPM-DistList.tnef").read_bytes()
    t = TNEF(tnef_data)

    assert str(t) == "<TNEF:0x1708, 1 attachments>"

    t2 = t.attachments[0].embed

    assert str(t2) == "<TNEF:0x1708>"


def test_bad_signature():
    tnef_data = Path(__file__).read_bytes()

    with pytest.raises(ValueError, match=r"Wrong TNEF signature: 0x\w+"):
        TNEF(tnef_data)


def test_tnef_htmlbody_hebrew_text():
    tnef_data = (HERE / "examples" / "iso-8859-8-i.tnef").read_bytes()
    t = TNEF(tnef_data)
    unicode_slice = t.htmlbody[1795:1799]
    assert unicode_slice == u'\u05d5\u05d5\u05d0\u05dc', "Unexpected val in htmlbody"
