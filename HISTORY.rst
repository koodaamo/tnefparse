
tnefparse 1.3.0 (2018-12-01)
=============================

- drop Python 2.6 & 3.3 support
- Python 2/3 compatibility fixes
- more tests & example files (jrideout)
- overall improved testing & start tracking coverage
- lots of parsing improvements (jrideout)
- turn some unnecessary warnings into debug messages
- add tnefparse -p | --path option for setting attachment extraction path
- support more MAPI (PidTag) properties (jrideout)
- support RTF body extraction (jrideout)
- support extracting top level object attributues in msgprops (jrideout)
- util.raw_mapi & tnefparse.parseFile functions will be deprecated after 1.3

tnefparse 1.2.3, 2018-11-14
============================

- misc. fixes

tnefparse 1.2.2, 2017
======================

- have `TNEF` init raise ValueError on invalid TNEF signature, rather than calling sys.exit()
- `parseFile` convenience function should not expect a `self` parameter, removed
- other misc. fixes

tnefparse 1.2.1, 2013
======================

- Python 3 compatibility; tests pass on Python 2.6/2.7/3.2/3.3
- add package to travis ci
- add tox.ini for testing using https://testrun.org/tox

tnefparse 1.2, 2013
===================

- performance improvements & bug fixes (Dave Baggett)
- added to_zip function for converting TNEF attachments into ZIPped ones (Dave Baggett)
- tnefparse is now used in the inky email client from Arcode

tnefparse 1.1.1, 08/2012 (unreleased)
=====================================

- fixed entry point bug that caused 'tnefparse' cmd-line invocation to fail

tnefparse 1.1, 03/2012
=======================

- Repackaged and renamed the library
- Code moved to github
- Use the stdlib logging module
- Further bug fixes and enhancements to pure-python code
- Add a command-line script
- Drop the unix tnef command-line tool wrapper

pytnef 0.2.1-Novell, circa 2010
================================

- Bug fixes/enhancements to pure-python code (Tom Doman)

pytnef 0.2, circa 2005
======================

- Added a wrapper for the unix tnef command-line tool (Petri Savolainen)
- Pure-python code not very useful yet

pytnef 0.1 - circa 2005
=======================

- First version (pure-python) created as a conversion of a Ruby TNEF decoder
  by Trevor Scheroeder (Petri Savolainen)
