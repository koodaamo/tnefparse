tnefparse - TNEF decoding and attachment extraction
===================================================

This is a pure-python library for decoding Microsoft's Transport Neutral Encapsulation Format (TNEF), for Python
versions 2.6 and above. Note that there are some issues with setuptools on Python 3.2, so when using Python3, version 3.3 or above is suggested.

For more information on TNEF, see for example `wikipedia <http://en.wikipedia.org/wiki/Transport_Neutral_Encapsulation_Format>`_. The full TNEF specification
is also available as a `PDF download <https://interoperability.blob.core.windows.net/files/MS-OXTNEF/[MS-OXTNEF].pdf>`_.

The library can be used as a basis for applications that need to parse TNEF. A command-line utility "tnefparse" is
also provided that gets installed as part of the package, that can be used to list contents of TNEF files and
extract attachments found inside them (requires python >= 2.7).

Use 'python setup.py test' or 'python runtests.py' to run the tests.

Issues and pull requests welcome. **Please however always provide an example TNEF file** that can be used to demonstrate the bug or desired behavior, if at all possible.

**Note: If you have understanding of TNEF and/or MIME internals or just need this package and want to help with maintaining it, I am open
to giving you commit rights. Just let me know.**

.. image:: https://badge.fury.io/py/tnefparse.png
    :target: http://badge.fury.io/py/tnefparse

.. image:: https://travis-ci.org/koodaamo/tnefparse.png?branch=master
        :target: https://travis-ci.org/koodaamo/tnefparse
