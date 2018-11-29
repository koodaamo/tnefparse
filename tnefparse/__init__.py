from .tnef import TNEF, TNEFAttachment, TNEFObject

def parseFile(fileobj):
   "a convenience function that returns a TNEF object"
   raise PendingDeprecationWarning("parseFile will be deprecated after 1.3")
   return TNEF(fileobj.read())
