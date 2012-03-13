from tnef import TNEF, TNEFAttachment, TNEFObject

def parseFile(self, fileobj):
   "a convenience function that returns a TNEF object"
   return TNEF(fileobj.read())
