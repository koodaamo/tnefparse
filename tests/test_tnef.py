import os
import pytest
from tnefparse.tnef import TNEF

THIS_DIR = os.path.dirname(os.path.abspath(__file__))


def test_tnef_str_representation():
    tnef_data = open(os.path.join(THIS_DIR, "examples", "two-files.tnef"), mode="rb").read()
    t = TNEF(tnef_data)

    assert str(t) == "<TNEF:0x237, 2 attachments>"


def test_tnef_msg_embed():
    tnef_data = open(os.path.join(THIS_DIR, "examples", "IPM-DistList.tnef"), mode="rb").read()
    t = TNEF(tnef_data)

    assert str(t) == "<TNEF:0x1708, 1 attachments>"

    t2 = t.attachments[0].embed

    assert str(t2) == "<TNEF:0x1708>"


def test_bad_signature():
    tnef_data = open(os.path.join(THIS_DIR, __file__), mode="rb").read()

    with pytest.raises(ValueError, match=r"Wrong TNEF signature: 0x\w+"):
        TNEF(tnef_data)
