from .tnef import TNEF, TNEFAttachment, TNEF_Object

def to_zip(data, default_name='no-name', deflate=True):
	"Convert attachments in TNEF data to zip format. Accepts and returns str type."
	# Parse the TNEF data
	tnef = TNEF(data)
	# Convert the TNEF file to an equivalent ZIP file
	tozip = {}
	for attachment in tnef.attachments:
		filename = attachment.name or default_name
		L = len(tozip.get(filename, []))
		if L > 0:
			# uniqify this file name by adding -<num> before the extension
			root, ext = os.path.splitext(filename)
			tozip[filename].append((attachment.data, "%s-%d%s" % (root, L + 1, ext)))
		else:
			tozip[filename] = [(attachment.data, filename)]
	# Add each attachment in the TNEF file to the zip file
	from zipfile import ZipFile, ZIP_DEFLATED, ZIP_STORED
	from cStringIO import StringIO
	sfp = StringIO()
	z = ZipFile(sfp, "w", ZIP_DEFLATED if deflate else ZIP_STORED)
	with z:
		for filename, entries in tozip.items():
			for entry in entries:
				data, name = entry
				z.writestr(name, data)
	# Return the binary data for the zip file
	return sfp.getvalue()

def from_file(self, fileobj):
	"a convenience function that returns a TNEF object"
	return TNEF(fileobj.read())
