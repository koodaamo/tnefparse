"""MAPI attribute definitions"""

import logging
from decimal import Decimal

from .util import (
    apptime,
    dbl64,
    float32,
    guid,
    int8,
    int16,
    int32,
    int64,
    systime,
    uint16,
    uint32,
)
from . import properties

logger = logging.getLogger(__package__)


SZMAPI_UNSPECIFIED    = 0x0000  # MAPI Unspecified  # noqa: E221
SZMAPI_NULL           = 0x0001  # MAPI null property  # noqa: E221
SZMAPI_SHORT          = 0x0002  # MAPI short (signed 16 bits)  # noqa: E221
SZMAPI_INT            = 0x0003  # MAPI integer (signed 32 bits)  # noqa: E221
SZMAPI_FLOAT          = 0x0004  # MAPI float (4 bytes)  # noqa: E221
SZMAPI_DOUBLE         = 0x0005  # MAPI double  # noqa: E221
SZMAPI_CURRENCY       = 0x0006  # MAPI currency (64 bits)  # noqa: E221
SZMAPI_APPTIME        = 0x0007  # MAPI application time  # noqa: E221
SZMAPI_ERROR          = 0x000A  # MAPI error (32 bits)  # noqa: E221
SZMAPI_BOOLEAN        = 0x000B  # MAPI boolean (16 bits)  # noqa: E221
SZMAPI_OBJECT         = 0x000D  # MAPI embedded object  # noqa: E221
SZMAPI_INT8BYTE       = 0x0014  # MAPI 8 byte signed int  # noqa: E221
SZMAPI_STRING         = 0x001E  # MAPI string  # noqa: E221
SZMAPI_UNICODE_STRING = 0x001F  # MAPI unicode-string (null terminated)  # noqa: E221
SZMAPI_SYSTIME        = 0x0040  # MAPI time (64 bits)  # noqa: E221
SZMAPI_CLSID          = 0x0048  # MAPI OLE GUID  # noqa: E221
SZMAPI_BINARY         = 0x0102  # MAPI binary  # noqa: E221

MULTI_VALUE_FLAG = 0x1000
GUID_EXISTS_FLAG = 0x8000

IMESSAGE_SIG = b"\x07\x03\x02\x00\x00\x00\x00\x00\xc0\x00\x00\x00\x00\x00\x00\x46"
IMESSAGE_SIG_LEN = len(IMESSAGE_SIG)


def decode_mapi(data, codepage='cp1252', starting_offset=None):
    """decode MAPI types"""

    data_len = len(data)
    attrs = []
    offset = starting_offset or 0
    num_properties = uint32(data, offset)
    offset += 4

    try:
        for i in range(num_properties):
            if offset >= data_len:
                logger.warning("Skipping property %r", i)
                continue

            attr_type = uint16(data, offset)
            offset += 2
            attr_name = uint16(data, offset)
            offset += 2

            # logger.debug("Attribute type: 0x%4.4x", attr_type)
            # logger.debug("Attribute name: 0x%4.4x", attr_name)
            guid_id = ''
            guid_name = None
            guid_prop = None
            if attr_name >= GUID_EXISTS_FLAG:
                guid_id = guid(data, offset)
                offset += 16
                kind = uint32(data, offset)
                offset += 4
                if kind == 0:
                    guid_prop = uint32(data, offset)
                    offset += 4
                else:
                    iid_len = uint32(data, offset)
                    offset += 4
                    r = iid_len % 4
                    if r != 0:
                        iid_len += 4 - r
                    guid_name = data[offset: offset + iid_len].decode('utf-16')
                    offset += iid_len

            num_mv_properties = None
            if MULTI_VALUE_FLAG & attr_type:
                attr_type ^= MULTI_VALUE_FLAG
                num_mv_properties = uint32(data, offset)
                offset += 4

            for mv in range(num_mv_properties or 1):
                try:
                    attr_data, offset = parse_property(data, offset, attr_name, attr_type, codepage, num_mv_properties)
                    attr = TNEFMAPI_Attribute(attr_type, attr_name, attr_data, guid_id, guid_name, guid_prop)
                    attrs.append(attr)
                    # print(attr)
                except Exception:
                    logger.debug("Attribute type: 0x%4.4x", attr_type)
                    logger.debug("Attribute name: 0x%4.4x (%s)", attr_name, TNEFMAPI_Attribute.codes.get(attr_name))
                    raise

            if (num_mv_properties or 1) % 2 and attr_type in (SZMAPI_SHORT, SZMAPI_BOOLEAN):
                offset += 2

    except Exception as exc:
        logger.error('decode_mapi exception: %s', exc)
        logger.debug('exception details:', exc_info=True)

    if starting_offset is not None:
        return (offset, attrs)
    else:
        return attrs


def parse_property(data, offset, attr_name, attr_type, codepage, is_multi):
    attr_data = None

    if attr_type in (SZMAPI_SHORT, SZMAPI_BOOLEAN):
        attr_data = data[offset: offset + 2]
        offset += 2
    elif attr_type in (SZMAPI_INT, SZMAPI_FLOAT, SZMAPI_ERROR):
        attr_data = data[offset: offset + 4]
        offset += 4
    elif attr_type in (SZMAPI_DOUBLE, SZMAPI_APPTIME, SZMAPI_CURRENCY, SZMAPI_INT8BYTE, SZMAPI_SYSTIME):
        attr_data = data[offset: offset + 8]
        offset += 8
    elif attr_type == SZMAPI_CLSID:
        attr_data = data[offset: offset + 16]
        offset += 16
    elif attr_type in (SZMAPI_STRING, SZMAPI_UNICODE_STRING, SZMAPI_OBJECT, SZMAPI_BINARY, SZMAPI_UNSPECIFIED):
        if is_multi:
            num_vals = 1
        else:
            num_vals = uint32(data, offset)
            offset += 4

        attr_data = []
        for j in range(num_vals):
            length = uint32(data, offset)
            offset += 4
            r = length % 4
            if r != 0:
                length += 4 - r

            item = data[offset: offset + length]
            if attr_type == SZMAPI_UNICODE_STRING:
                item = item.decode('utf-16')
            elif attr_type == SZMAPI_STRING:
                item = item.decode(codepage)
            attr_data.append(item)
            offset += length

    else:
        raise ValueError(f"Unknown MAPI type {attr_type:#06x}")

    return attr_data, offset


class TNEFMAPI_Attribute:
    """represents a mapi attribute

    Property reference docs:

    https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/mapping-canonical-property-names-to-mapi-names#tagged-properties
    https://fossies.org/linux/libpst/xml/MAPI_definitions.pdf

    Most these properties represent PidTag properties
    """
    codes = properties.CODE_TO_NAME

    OutlookGuid = '05133f00aa00da98101b450b6ed8da90'
    AppointmentGuid = '46000000000000c00000000000062002'

    def __init__(self, attr_type, name, data, guid_id, guid_name=None, guid_prop=None):
        self.attr_type = attr_type
        self.name = name
        self.raw_data = data
        self.guid = guid_id
        self.guid_name = guid_name
        self.guid_prop = guid_prop
        if self.guid_name:
            self.guid_name = self.guid_name.rstrip('\x00')

    @property
    def data(self):
        if self.attr_type == SZMAPI_NULL:
            return None
        elif self.attr_type == SZMAPI_SHORT:
            return int16(self.raw_data)
        elif self.attr_type == SZMAPI_INT:
            return int32(self.raw_data)
        elif self.attr_type == SZMAPI_FLOAT:
            return float32(self.raw_data)
        elif self.attr_type == SZMAPI_DOUBLE:
            return dbl64(self.raw_data)
        elif self.attr_type == SZMAPI_CURRENCY:
            return Decimal(int64(self.raw_data)) / Decimal(10000)
        elif self.attr_type == SZMAPI_APPTIME:
            return apptime(self.raw_data)
        elif self.attr_type == SZMAPI_ERROR:
            return uint32(self.raw_data)
        elif self.attr_type == SZMAPI_BOOLEAN:
            return bool(uint16(self.raw_data))
        elif self.attr_type == SZMAPI_INT8BYTE:
            return int8(self.raw_data)
        elif self.attr_type == SZMAPI_SYSTIME:
            return systime(self.raw_data)
        elif self.attr_type == SZMAPI_CLSID:
            return guid(self.raw_data)
        elif self.attr_type in (SZMAPI_STRING, SZMAPI_UNICODE_STRING):
            return ''.join([s.rstrip('\x00') for s in self.raw_data])

        if self.attr_type not in (SZMAPI_UNSPECIFIED, SZMAPI_BINARY, SZMAPI_OBJECT):
            logger.warning(f"Unknown data type: 0x{self.attr_type:04x}")

        return b''.join([s.rstrip(b'\x00') for s in self.raw_data])

    @property
    def name_str(self):
        return (
            self.guid_name or
            TNEFMAPI_Attribute.codes.get(self.name) or
            TNEFMAPI_Attribute.codes.get(self.guid_prop) or
            hex(self.guid_prop or self.name)
        )

    def __str__(self):
        return f"<ATTR: {self.name_str}>"
