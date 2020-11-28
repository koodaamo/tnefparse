import os
from tnefparse.tnef import TNEF

THIS_DIR = os.path.dirname(os.path.abspath(__file__))


def test_tnef_str_representation():
    tnef_data = open(os.path.join(THIS_DIR, "examples", "two-files.tnef"), mode="rb").read()
    t = TNEF(tnef_data)

    assert str(t) == "<TNEF:0x237, 2 attachments>"
