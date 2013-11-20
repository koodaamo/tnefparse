"""utility functions
"""
import sys
import logging
logger = logging.getLogger("tnef-decode")


# For compatibility, added two versions of bytes_to_int and checksum for now.
# TODO: refactor?

def bytes_to_int_py3(bytes=None):
   "transform multi-byte values into integers, python3 version"
   # NOTE: byte ordering & signage based on trial and error against tests
   return int.from_bytes(bytes, byteorder="little", signed=False)

def bytes_to_int_py2(bytes=None):
   "transform multi-byte values into integers, python2 version"
   n = num = 0
   for b in bytes:
      num += ( ord(b) << n)
      n += 8
   return num

def checksum_py3(data):
   "calculate a checksum for the TNEF data"
   # TODO: use int.from_bytes
   return sum([x for x in data]) & 0xFFFF

def checksum_py2(data):
   "calculate a checksum for the TNEF data"
   # NOTE: under Python 3.2+ you can do this much faster with int.from_bytes
   return sum([ord(x) for x in data]) & 0xFFFF


if sys.hexversion > 0x03000000:
  bytes_to_int = bytes_to_int_py3
  checksum = checksum_py3
else:
  bytes_to_int = bytes_to_int_py2
  checksum = checksum_py2


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
