"""utility functions
"""
import logging
import struct
import uuid
from datetime import datetime, timedelta

logger = logging.getLogger(__package__)


def make_unpack(structure):
    call = struct.Struct(structure).unpack_from

    def unpack(byte_arr, offset=0):
        return call(byte_arr, offset)[0]

    return unpack


uint8 = make_unpack('<B')
int8 = make_unpack('<b')
uint16 = make_unpack('<H')
int16 = make_unpack('<h')
uint32 = make_unpack('<I')
int32 = make_unpack('<i')
uint64 = make_unpack('<Q')
int64 = make_unpack('<q')
float32 = make_unpack('<f')
dbl64 = make_unpack('<d')


def guid(byte_arr, offset=0):
    return uuid.UUID(bytes_le=byte_arr[offset: offset + 16])


OLE_TIME_ZERO = datetime(1899, 12, 30)
EPOCH_AS_FILETIME = 116444736000000000  # January 1, 1970 as MS file time
HUNDREDS_OF_NANOSECONDS = 10000000


def systime(byte_arr, offset=0):
    ft = uint64(byte_arr, offset)
    try:
        return datetime.utcfromtimestamp((ft - EPOCH_AS_FILETIME) / HUNDREDS_OF_NANOSECONDS)
    except:  # noqa: E722
        microseconds = ft / 10
        return datetime(1601, 1, 1) + timedelta(microseconds=microseconds)


def apptime(byte_arr, offset=0):
    return timedelta(dbl64(byte_arr, offset)) + OLE_TIME_ZERO


dt_parts = struct.Struct('<HHHHHH').unpack_from


def typtime(byte_arr, offset=0):
    parts = dt_parts(byte_arr, offset)
    return datetime(*parts)


def bytes_to_int(byte_arr):
    """transform multi-byte values into integers"""
    return int.from_bytes(byte_arr, byteorder="little", signed=False)


def checksum(data):
    return sum(bytearray(data)) & 0xFFFF
