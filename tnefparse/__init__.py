import warnings
from .tnef import TNEF, TNEFAttachment, TNEFObject


def parseFile(fileobj):
    "a convenience function that returns a TNEF object"
    warnings.warn("parseFile will be deprecated after 1.3", DeprecationWarning)
    return TNEF(fileobj.read())
