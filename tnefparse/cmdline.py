import sys, argparse, logging

logging.basicConfig()
logging.root.setLevel(logging.ERROR)

from tnef import TNEF, TNEFAttachment, TNEFObject
from mapi import TNEFMAPI_Attribute


descr = 'Extract TNEF file contents. Show this help message if no arguments are given.'
parser = argparse.ArgumentParser(description=descr)
argument = parser.add_argument

argument('file', type=argparse.FileType('rb'), nargs='+', action="append",
         help='space-separated list of paths to the TNEF files')

argument('-o','--overview', action='store_true',
         help='show long debug overview of TNEF file contents')

argument('-a','--attachments', action='store_true',
         help='extract attachments to current directory')

argument('-b', '--body', action='store_true', help='extract the body to stdout')

argument('-hb', '--htmlbody', action='store_true', help='extract the HTML body to stdout')

argument('-l', '--logging', help="enable logging by setting a log level",
         choices = ["DEBUG", "INFO", "WARN", "ERROR"])

argument('-c', '--checksum', help="calculate checksums (off by default)",
         action="store_true", default=False)



def tnefparse():
   "command-line script"

   if len(sys.argv)==1:
       parser.print_help()
       sys.exit(1)

   args = parser.parse_args()

   if args.logging:
      level = eval("logging." + args.logging)
      logging.root.setLevel(level)

   for tfp in args.file[0]:
      t = TNEF(tfp.read(), do_checksum=args.checksum)
      if args.overview:
         print("\nOverview of %s: \n" % tfp.name)

         # list TNEF attachments
         print("  Attachments:\n")
         for a in t.attachments:
            print("    " + a.name)

         # list TNEF objects
         print("\n  Objects:\n")
         print("    " + "\n    ".join([TNEF.codes[o.name] for o in t.objects]))

         # list TNEF MAPI properties
         print("\n  Properties:\n")
         for p in t.mapiprops:
            try:
               print("    " + TNEFMAPI_Attribute.codes[p.name])
            except KeyError:
               logging.root.warn("Unknown MAPI Property: %s" % hex(p.name))
         print("")

      elif args.attachments:
         for a in t.attachments:
            with open(a.name, "wb") as afp:
               afp.write(a.data)
         sys.exit("Successfully wrote %i files" % len(t.attachments))

      if args.body:
         print(getattr(t, "body", "No body found"))

      if args.htmlbody:
         body = getattr(t, "htmlbody", ["No HTML body found"])
         print(body[0])
