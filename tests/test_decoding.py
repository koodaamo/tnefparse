import os, sys, logging, tempfile

logging.basicConfig()

from tnefparse import TNEF
from tnefparse.tnef import to_zip

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
   ),
   ("MAPI_ATTACH_DATA_OBJ.tnef", 0x1af, ['VIA_Nytt_1402.doc', 'VIA_Nytt_1402.pdf', 'VIA_Nytt_14021.htm'],
       [0x9006, 0x9007, 0x9003, 0x9002, 0x9005, 0x9002, 0x9005, 0x9002, 0x9005]),
   ("MAPI_OBJECT.tnef", 0x08, ['Untitled_Attachment'], []),
   ("garbage-at-end.tnef", 0x415, [], []),
   ("long-filename.tnef", 0x1422, ['allproductsmar2000.dat'], []),
   ("missing-filenames.tnef", 0x601, ['generpts.src', 'TechlibDEC99.doc', 'TechlibDEC99-JAN00.doc', 'TechlibNOV99.doc'], []),
   ("multi-value-attribute.tnef", 0x1512, ['208225__5_seconds__Voice_Mail.mp3'], []),
   ("one-file.tnef", 0x237, ['AUTHORS'], []),
   ("rtf.tnef", 0xc02, [], []),
   ("triples.tnef", 0xea64, [], []),
   ("unicode-mapi-attr-name.tnef", 0x69ec, ['spaconsole2.cfg', 'image001.png', 'image002.png', 'image003.png'], []),
   ("unicode-mapi-attr.tnef", 0x408f, ['example.dat'], []),
   ("tzanke.tnef", 0xa2e, ['TBZ PARIV GmbH.jpg', 'image003.jpg', u'UmlautAnhang-\xe4\xfc\xf6.txt'], []),
)

# generate tests for all example files
def pytest_generate_tests(metafunc):
    if "tnefspec" in metafunc.funcargnames:
        metafunc.parametrize("tnefspec", SPECS, ids=[s[0] for s in SPECS])


objnames = lambda t: [TNEF.codes[o.name] for o in t.objects]
objcodes = lambda t: [o.name for o in t.objects]

def test_decode(tnefspec):
   fn, key, attchs, objs = tnefspec
   with open(datadir + os.sep + fn, "rb") as tfile:
      t = TNEF(tfile.read())
      assert t.key == key, "wrong key: 0x%2.2x" % t.key
      assert [a.long_filename() for a in t.attachments] == attchs
      for m in t.mapiprops:
          assert m.data is not None
      if t.htmlbody:
          assert b'html' in t.htmlbody

      # assert [a.name.decode() for a in t.attachments] == [to_shortname(a) for a in attchs]

      # assert objcodes(t) == objs, "wrong objs: %s" % ["0x%2.2x" % o.name for o in t.objects]

def test_zip():
    with open(datadir + os.sep + 'one-file.tnef', "rb") as tfile:
        zip_data = to_zip(tfile.read())
        with tempfile.TemporaryFile() as out:
            out.write(zip_data)


def to_shortname(longname):
    if len(longname) < 15:
        return longname
    root, ext = os.path.splitext(longname)
    strip = len(ext)
    return root.upper()[0:10-strip] + '~1' + ext.upper()
