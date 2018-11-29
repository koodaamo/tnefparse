tnefparse - TNEF decoding and attachment extraction
===================================================

This is a pure-python library for decoding Microsoft's Transport Neutral Encapsulation Format (TNEF), for Python
versions 2.7, 3.5+ and PyPy.

For more information on TNEF, see for example `wikipedia <http://en.wikipedia.org/wiki/Transport_Neutral_Encapsulation_Format>`_. The full TNEF specification
is also available as a `PDF download <https://interoperability.blob.core.windows.net/files/MS-OXTNEF/[MS-OXTNEF].pdf>`_.

The library can be used as a basis for applications that need to parse TNEF. To parse a file into a TNEF object, run eg. :

 >>> from tnefparse import TNEF
 >>> with open("tests/examples/one-file.tnef", "rb") as tneffile:
 ...    tnefobj = TNEF(tneffile.read())

A :code:`tnefparse` command-line utility is also provided for listing contents of TNEF files, extracting attachments found inside them and so on.

Use :code:`python setup.py test` or :code:`python runtests.py` to run the tests.

Issues and pull requests welcome. **Please however always provide an example TNEF file** that can be used to demonstrate the bug or desired behavior, if at all possible.

**Note: If you have understanding of TNEF and/or MIME internals or just need this package and want to help with maintaining it, I am open to giving you commit rights. Just let me know.**

.. image:: https://badge.fury.io/py/tnefparse.png
    :target: http://badge.fury.io/py/tnefparse

.. image:: https://travis-ci.org/koodaamo/tnefparse.png?branch=master
        :target: https://travis-ci.org/koodaamo/tnefparse

.. highlight:: python
