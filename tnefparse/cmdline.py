import argparse
import json
import logging
import os
import sys

from . import properties
from .tnef import TNEF, to_zip

logging.basicConfig()
logger = logging.getLogger(__package__)
logger.setLevel(logging.ERROR)


descr = 'Extract TNEF file contents. Show this help message if no arguments are given.'
parser = argparse.ArgumentParser(description=descr)
argument = parser.add_argument

argument('file', type=argparse.FileType('rb'), nargs='+', action="append",
         help='space-separated list of paths to the TNEF files')

argument('-o', '--overview', action='store_true',
         help='show (possibly long) overview of TNEF file contents')

argument('-a', '--attachments', action='store_true',
         help='extract attachments, by default to current dir')

argument('-z', '--zip', action='store_true',
         help='extract attachments into a single zip file, by default to current dir')

argument('-p', '--path',
         help='optional explicit path to extract attachments to')

argument('-b', '--body', action='store_true',
         help='extract the body to stdout')

argument('-hb', '--htmlbody', action='store_true',
         help='extract the HTML body to stdout')

argument('-rb', '--rtfbody', action='store_true',
         help='extract the RTF body to stdout')

argument('-l', '--logging', choices=["DEBUG", "INFO", "WARN", "ERROR"],
         help="enable logging by setting a log level")

argument('-c', '--checksum', action="store_true", default=False,
         help="calculate checksums (off by default)")

argument('-d', '--dump', action="store_true", default=False,
         help="extract a json dump of the tnef contents")


def tnefparse(argv=None) -> None:
    """command-line script"""

    if argv is None:
        argv = sys.argv[1:]

    if not argv:
        parser.print_help()
        sys.exit(1)

    args = parser.parse_args(argv)

    if args.logging:
        logger.setLevel(getattr(logging, args.logging))

    for tfp in args.file[0]:
        try:
            t = TNEF(tfp.read(), do_checksum=args.checksum)
        except ValueError as exc:
            sys.exit(str(exc))
        if args.overview:
            print(f"\nOverview of {tfp.name}: \n")

            # list TNEF attachments
            print("  Attachments:\n")
            for a in t.attachments:
                print("    " + a.long_filename())

            # list TNEF objects
            print("\n  Objects:\n")
            print("    " + "\n    ".join([TNEF.codes[o.name] for o in t.objects]))

            # list TNEF MAPI properties
            print("\n  Properties:\n")
            for p in t.mapiprops:
                try:
                    print("    " + properties.CODE_TO_NAME[p.name])
                except KeyError:
                    logger.warning("Unknown MAPI Property: 0x%x", p.name)
            print("")

        elif args.dump:
            print(json.dumps(t.dump(force_strings=True), sort_keys=True, indent=4))

        elif args.attachments:
            pth = args.path.rstrip(os.sep) + os.sep if args.path else ''
            for a in t.attachments:
                with open(pth + a.long_filename(), "wb") as afp:
                    afp.write(a.data)
            sys.stderr.write(f"Successfully wrote {len(t.attachments)} files\n")
            sys.exit()

        elif args.zip:
            zipped = to_zip(t)
            pth = args.path.rstrip(os.sep) + os.sep if args.path else ''
            with open(pth + 'attachments.zip', "wb") as afp:
                afp.write(zipped)
            sys.stderr.write("Successfully wrote attachments.zip\n")
            sys.exit()

        def print_body(attr: str, description: str) -> None:
            body = getattr(t, attr)
            if body is None:
                sys.exit(f"No {description} found")
            elif isinstance(body, bytes):
                sys.stdout.write(body.decode('latin-1'))
            else:
                sys.stdout.write(body)

        if args.body:
            print_body("body", "body")
        if args.htmlbody:
            print_body("htmlbody", "HTML body")
        if args.rtfbody:
            print_body("rtfbody", "RTF body")
