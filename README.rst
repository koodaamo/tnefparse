tnefparse - TNEF decoding and attachment extraction
===================================================

This is a pure-python library for decoding Microsoft's Transport Neutral Encapsulation Format (TNEF), for Python
versions 2.6 and above. Note that there are some issues with setuptools on Python 3.2, so when using Python3, version 3.3 or above is suggested.

See for example wikipedia for `more on TNEF <http://en.wikipedia.org/wiki/Transport_Neutral_Encapsulation_Format>`_.

The library can be used as a basis for applications that need to parse TNEF. A command-line utility "tnefparse" is
also provided that gets installed as part of the package, that can be used to list contents of TNEF files and
extract attachments found inside them (requires python >= 2.7).

Use 'python setup.py test' or 'python runtests.py' to run the tests.

Please see the HISTORY.txt file for full release notes.

Issues and pull requests welcome. **Please however always provide an example TNEF file** that can be used to demonstrate the bug or desired behavior, if at all possible.

**If you have understanding of TNEF and/or MIME internals and want to help with maintaining this package, file an issue or email me. I'd be happy to give you commit rights.**


.. image:: https://badge.fury.io/py/tnefparse.png
    :target: http://badge.fury.io/py/tnefparse

.. image:: https://travis-ci.org/koodaamo/tnefparse.png?branch=master
        :target: https://travis-ci.org/koodaamo/tnefparse
