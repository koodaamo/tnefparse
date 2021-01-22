from datetime import datetime
from decimal import Decimal
from uuid import UUID
import logging

from tnefparse.mapi import TNEFMAPI_Attribute

attributes = [
    (0x0000, [], b""),
    (0x0001, [], None),
    (0x0002, b"\00\01", 256),
    (0x0003, b"4\x01\x00\x00", 308),
    (0x0003, b"\x9fN\x00\x00", 20127),
    (0x0003, b"\xff\xff\xff\xff", -1),
    (0x0004, b"ABCD", 781.0352172851562),
    (0x0005, b"ABCDEFGH", 1.5839800103804824e+40),
    (0x0006, b"ABCDEFGH", Decimal('520820875738921.4273')),
    (0x0007, b"0\xe3\x1d\x0e\xef\x15\xbf\x01", datetime(1899, 12, 30, 0, 0)),
    (0x000A, b"\xff\xff\xff\xff", 4294967295),
    (0x000B, b"\x00\x00", False),
    (0x000B, b"\x01\x00", True),
    (0x000D, [b"00"], b"00"),
    (0x0014, b"\x01\x00", 1),
    (0x001E, ["<14341.17573.560761.368512@localhost.localdomain>\x00\x00\x00"],
        "<14341.17573.560761.368512@localhost.localdomain>"),
    (0x001E, ["AUTHORS\x00"], "AUTHORS"),
    (0x001F, ["\u200bhello world\x00\x00"], "\u200bhello world"),
    (0x001F, ["example.dat\x00"], "example.dat"),
    (0x0040, b"0\xe3\x1d\x0e\xef\x15\xbf\x01", datetime(1999, 10, 14, 2, 51, 46, 787000)),
    (0x0048, b"\x11\xf8\x30\x73\x7f\xf4\xbc\x41\xa4\xff\xe7\x92\xd0\x73\xf4\x1f",
        UUID('7330f811-f47f-41bc-a4ff-e792d073f41f')),
    (0x0102, [b""], b""),
    (0x0102, [b"<14341.17573.560761.368512@localhost.localdomain>\x00\x00\x00"],
        b"<14341.17573.560761.368512@localhost.localdomain>"),
    (0x0102, [b"EX:/O=COMPUWARE/OU=NUMEGA LAB/CN=RECIPIENTS/CN=MSIMPSON\x00"],
        b"EX:/O=COMPUWARE/OU=NUMEGA LAB/CN=RECIPIENTS/CN=MSIMPSON"),
]


def test_decode_known(caplog):
    for typ, data, parsed in attributes:
        att = TNEFMAPI_Attribute(typ, 0, data, None, f"Type: 0x{typ:04x}")
        assert att.data == parsed, att.name_str

    assert caplog.record_tuples == []


def test_decode_unknown(caplog):
    att = TNEFMAPI_Attribute(0xAAAA, 0, [b"123"], None)
    assert att.data == b'123'

    assert caplog.record_tuples == [
        ('tnefparse', logging.WARNING, 'Unknown data type: 0xaaaa'),
    ]
