# https://docs.microsoft.com/en-us/windows/desktop/intl/code-page-identifiers
CODEPAGE_MAP = {
    20127: 'ascii',
    20866: 'koi8-r',
    28591: 'iso-8859-1',
    65001: 'utf-8',
}
FALLBACK = 'cp1252'


class Codepage(object):
    def __init__(self, codepage):
        self.cp = codepage

    def best_guess(self):
        if CODEPAGE_MAP.get(self.cp):
            return CODEPAGE_MAP.get(self.cp)
        elif self.cp <= 1258:  # max cpXXXX page in python
            return 'cp%d' % self.cp
        else:
            return None

    def codepage(self):
        bg = self.best_guess()
        if bg:
            return bg
        else:
            return 'cp%d' % self.cp

    def decode(self, byte_str):
        if isinstance(byte_str, bytes):
            return byte_str.decode(self.best_guess() or FALLBACK)
        else:
            return byte_str
