import os, sys, logging

logging.basicConfig()

from tnefparse import TNEF

datadir = os.path.dirname(os.path.realpath(__file__)) + os.sep + "examples"

SPECS = (
   ("body.tnef", 0x1125, [],
     [0x9006, 0x9007, 0x800d, 0x8005, 0x8020, 0x8009, 0x9004, 0x9003]
   ),

   ("two-files.tnef", 0x237, ["AUTHORS", "README"],
      [0x9006, 0x9007, 0x8008, 0x8009, 0x06, 0x8020, 0x8005, 0x8004, 0x800d, 0x9003,
      0x9002, 0x8012, 0x8013, 0x8010, 0x800f, 0x9005, 0x9002, 0x8012, 0x8013, 0x8010,
      0x800f, 0x9005]
   ),
   ("data-before-name.tnef", 0x0d, ["AUTOEXEC.BAT", "CONFIG.SYS", "boot.ini"],
      [0x9006, 0x9007, 0x8008, 0x800d, 0x8006, 0x9003, 0x9002, 0x8013, 0x800f, 0x8010,
      0x8011, 0x9005, 0x9002, 0x8013, 0x800f, 0x8010, 0x8011, 0x9005, 0x9002, 0x8013,
      0x800f, 0x8010, 0x8011, 0x9005]
   ),
   ("multi-name-property.tnef", 0xc6c7, [],
      [0x9006, 0x9007, 0x9003]
   )
)


# generate tests for all example files
def pytest_generate_tests(metafunc):
    if "tnefspec" in metafunc.funcargnames:
        metafunc.parametrize("tnefspec", SPECS)


objnames = lambda t: [TNEF.codes[o.name] for o in t.objects]
objcodes = lambda t: [o.name for o in t.objects]

def test_decode(tnefspec):
   fn, key, attchs, objs = tnefspec
   with open(datadir + os.sep + fn, "rb") as tfile:
      t = TNEF(tfile.read())
      assert t.key == key, "wrong key: 0x%2.2x" % t.key
      assert objcodes(t) == objs, "wrong objs: %s" % ["0x%2.2x" % o.name for o in t.objects]
      assert [a.name for a in t.attachments] == attchs

