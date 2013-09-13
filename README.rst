yatp - Yet Another TNEF Parser 
==============================

This is a fork of tnefparse a pure-python library for decoding
Microsoft's Transport Neutral Encapsulation Format (TNEF).

The library can be used as a basis for applications that need to parse TNEF.
There are some tests; run 'python setup.py test' or 'python runtests.py'.

A command-line utility "tnefparse" is also provided that gets installed 
as part of the package, that can be used to list contents of TNEF files 
and extract attachments found inside them (requires python >= 2.7).

