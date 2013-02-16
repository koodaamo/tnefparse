tnefparse - TNEF decoding and attachment extraction
===================================================

This is a pure-python library for decoding Microsoft's Transport Neutral Encapsulation Format (TNEF).
See for example wikipedia for `more on TNEF <http://en.wikipedia.org/wiki/Transport_Neutral_Encapsulation_Format>`_.

The library can be used as a basis for applications that need to parse TNEF.
There are some tests; run 'python setup.py test' or 'python runtests.py'.

A command-line utility "tnefparse" is also provided that gets installed as part of the package, 
that can be used to list contents of TNEF files and extract attachments found inside them (requires python >= 2.7).

What's new in this release:

 - performance enhancements and bug fixes
 - new to_zip function for converting TNEF attachments into ZIPped ones

Contributions in this release mostly by Dave Baggett of Arcode, makers of the inky email client.

Please see the HISTORY.txt file for full release notes.
