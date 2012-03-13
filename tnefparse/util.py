"""utility functions
"""
import logging
logger = logging.getLogger("tnef-decode")

def bytes_to_int(bytes=None):
   "transform multi-byte values into integers"
   n = num = 0
   for b in bytes:
      num += ( ord(b) << n)
      n += 8
   return num


def checksum(data):
   "calculate a checksum for the TNEF data"
   sum = 0
   for byte in data:
      sum = (sum + ord(byte)) % 65536
   return sum


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
