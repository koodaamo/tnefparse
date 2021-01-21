from tnefparse.codepage import Codepage


def test_guesses():
    codes = [0, 437, 1200, 1252, 1255, 20127, 20866, 28591, 57011, 65001, 65002]
    best_guesses = [Codepage(c).best_guess() for c in codes]
    assert best_guesses == ['cp0', 'cp437', 'cp1200', 'cp1252', 'cp1255', 'ascii',
                            'koi8-r', 'iso-8859-1', None, 'utf-8', None]


def test_codepage():
    codes = [20127, 1252, 57011]
    pages = [Codepage(c).codepage() for c in codes]
    assert pages == ['ascii', 'cp1252', 'cp57011']


def test_decode_str():
    cp = Codepage(57011)
    assert '123' == cp.decode('123')


def test_decode_bg():
    cp = Codepage(65001)
    assert "☺" == cp.decode(b"\xE2\x98\xBA")


def test_decode_fallback():
    cp = Codepage(9999)
    assert 'ñ' == cp.decode(b'\xF1')
