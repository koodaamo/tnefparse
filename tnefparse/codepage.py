from typing import Optional, Union

# https://docs.microsoft.com/en-us/windows/desktop/intl/code-page-identifiers
CODEPAGE_MAP = {
    20127: 'ascii',
    20866: 'koi8-r',
    28591: 'iso-8859-1',
    65001: 'utf-8',
}
FALLBACK = 'cp1252'


class Codepage:
    def __init__(self, codepage: int) -> None:
        self.cp = codepage

    def best_guess(self) -> Optional[str]:
        if CODEPAGE_MAP.get(self.cp):
            return CODEPAGE_MAP.get(self.cp)
        elif self.cp <= 1258:  # max cpXXXX page in python
            return f"cp{self.cp}"
        else:
            return None

    def codepage(self) -> str:
        bg = self.best_guess()
        if bg:
            return bg
        else:
            return f"cp{self.cp}"

    def decode(self, byte_str: Union[str, bytes]) -> str:
        if isinstance(byte_str, bytes):
            return byte_str.decode(self.best_guess() or FALLBACK)
        else:
            return byte_str
