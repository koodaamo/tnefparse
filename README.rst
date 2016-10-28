tnefparse - TNEF decoding and attachment extraction
===================================================

This is a pure-python library for decoding Microsoft's Transport Neutral Encapsulation Format (TNEF), for Python
versions 2.6 and above.

See for example wikipedia for `more on TNEF <http://en.wikipedia.org/wiki/Transport_Neutral_Encapsulation_Format>`_.

The library can be used as a basis for applications that need to parse TNEF. A command-line utility "tnefparse" is
also provided that gets installed as part of the package, that can be used to list contents of TNEF files and
extract attachments found inside them (requires python >= 2.7).

Use 'python setup.py test' or 'python runtests.py' to run the tests.

Please see the HISTORY.txt file for full release notes.

.. image:: https://badge.fury.io/py/tnefparse.png
    :target: http://badge.fury.io/py/tnefparse

.. image:: https://travis-ci.org/koodaamo/tnefparse.png?branch=master
        :target: https://travis-ci.org/koodaamo/tnefparse
