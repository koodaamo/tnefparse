"""utility functions
"""
import logging
import struct
import sys
import uuid
import warnings
from datetime import datetime, timedelta

if sys.hexversion < 0x03000000:
    range = xrange

logger = logging.getLogger("tnef-decode")


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
    return uuid.UUID(bytes_le=byte_arr[offset : offset + 16])


OLE_TIME_ZERO = datetime(1899, 12, 30)
EPOCH_AS_FILETIME = 116444736000000000  # January 1, 1970 as MS file time
HUNDREDS_OF_NANOSECONDS = 10000000


def systime(byte_arr, offset=0):
    ft = uint64(byte_arr, offset)
    return datetime.utcfromtimestamp((ft - EPOCH_AS_FILETIME) / HUNDREDS_OF_NANOSECONDS)


def apptime(byte_arr, offset=0):
    return timedelta(dbl64(byte_arr, offset)) + OLE_TIME_ZERO


dt_parts = struct.Struct('<HHHHHH').unpack_from


def typtime(byte_arr, offset=0):
    parts = dt_parts(byte_arr, offset)
    return datetime(*parts)


def bytes_to_int_py3(byte_arr):
    "transform multi-byte values into integers, python3 version"
    return int.from_bytes(byte_arr, byteorder="little", signed=False)


def bytes_to_int_py2(byte_arr):
    "transform multi-byte values into integers, python2 version"
    n = num = 0
    for b in byte_arr:
        num += ord(b) << n
        n += 8
    return num


if sys.hexversion > 0x03000000:
    bytes_to_int = bytes_to_int_py3
else:
    bytes_to_int = bytes_to_int_py2


def checksum(data):
    return sum(bytearray(data)) & 0xFFFF


def raw_mapi(dataLen, data):
    "debug raw MAPI data when decoding MAPI types"
    warnings.warn("raw_mapi will be deprecated after 1.3", DeprecationWarning)
    loop = 0
    logger.debug("Raw MAPI Data:")
    while loop <= dataLen:
        if (loop + 16) < dataLen:
            logger.debug(
                "%2.2x%2.2x %2.2x%2.2x %2.2x%2.2x %2.2x%2.2x %2.2x%2.2x %2.2x%2.2x %2.2x%2.2x %2.2x%2.2x"
                % (
                    ord(data[loop]),
                    ord(data[loop + 1]),
                    ord(data[loop + 2]),
                    ord(data[loop + 3]),
                    ord(data[loop + 4]),
                    ord(data[loop + 5]),
                    ord(data[loop + 6]),
                    ord(data[loop + 7]),
                    ord(data[loop + 8]),
                    ord(data[loop + 9]),
                    ord(data[loop + 10]),
                    ord(data[loop + 11]),
                    ord(data[loop + 12]),
                    ord(data[loop + 13]),
                    ord(data[loop + 14]),
                    ord(data[loop + 15]),
                )
            )
        loop += 16
    loop -= 16
    q, r = divmod(dataLen, 16)
    strList = []
    for i in range(r):
        subq, subr = divmod(i, 2)
        if i != 0 and subr == 0:
            strList.append(' ')
        strList.append('%2.2x' % ord(data[loop + i]))
    logger.debug('%s' % ''.join(strList))
