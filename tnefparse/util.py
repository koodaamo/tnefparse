"""utility functions
"""
import sys
import struct
import logging
from datetime import datetime

if sys.hexversion < 0x03000000:
   range = xrange

logger = logging.getLogger("tnef-decode")


_unint32_unpack = struct.Struct('<I').unpack
def uint32(byte_arr):
    return _unint32_unpack(byte_arr)[0]


_unint16_unpack = struct.Struct('<H').unpack
def uint16(byte_arr):
    return _unint16_unpack(byte_arr)[0]


def bytes_to_int_py3(byte_arr):
   "transform multi-byte values into integers, python3 version"
   return int.from_bytes(byte_arr, byteorder="little", signed=False)


def bytes_to_int_py2(byte_arr):
   "transform multi-byte values into integers, python2 version"
   n = num = 0
   for b in byte_arr:
      num += ( ord(b) << n)
      n += 8
   return num


if sys.hexversion > 0x03000000:
  bytes_to_int = bytes_to_int_py3
else:
  bytes_to_int = bytes_to_int_py2


def checksum(data):
    return sum(bytearray(data)) & 0xFFFF


def bytes_to_date(byte_arr):
   return datetime(*[uint16(byte_arr[n:n+2]) for n in range(0, 12, 2)])


def raw_mapi(dataLen, data):
   "debug raw MAPI data when decoding MAPI types"
   loop = 0
   logger.debug("Raw MAPI Data:")
   while loop <= dataLen:
      if (loop + 16) < dataLen:
         logger.debug("%2.2x%2.2x %2.2x%2.2x %2.2x%2.2x %2.2x%2.2x %2.2x%2.2x %2.2x%2.2x %2.2x%2.2x %2.2x%2.2x" %
                (ord(data[loop]), ord(data[loop+1]), ord(data[loop+2]), ord(data[loop+3]),
                 ord(data[loop+4]), ord(data[loop+5]), ord(data[loop+6]), ord(data[loop+7]),
                 ord(data[loop+8]), ord(data[loop+9]), ord(data[loop+10]), ord(data[loop+11]),
                 ord(data[loop+12]), ord(data[loop+13]), ord(data[loop+14]), ord(data[loop+15])))
      loop += 16
   loop -= 16
   q,r = divmod(dataLen, 16)
   strList = []
   for i in range(r):
      subq,subr = divmod(i, 2)
      if i !=0 and subr == 0:
         strList.append(' ')
      strList.append('%2.2x' % ord(data[loop+i]))
   logger.debug('%s' % ''.join(strList))
