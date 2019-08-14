#!/usr/bin/python3.7

import sys, re, email, tempfile
import olefile
from email.parser import BytesParser
from email import policy

padding = '.'
placeholder = b'{{.RidPlaceholder}}'

EXT_RE = re.compile(r'(\.doc|\.xls|\.ppt)$', re.IGNORECASE)
RID_RE = re.compile(r'<!--\040RID:\040([a-z0-9]{6,})\040-->', re.IGNORECASE)

msg = email.message_from_file(sys.stdin, policy=policy.default)

for part in msg.iter_parts():
	filename = part.get_filename() # Get filename of part

	if not filename: # Body will not have a filename
		content = part.get_content()
		matches = RID_RE.search(content) # Search for RID string in content
		if matches:
			rid = matches.group(1)
			new_content = RID_RE.sub('', content)
			part.set_content(new_content)
			continue
		else:
			quit(1) # No matches, this isn't supposed to happen
	if EXT_RE.search(filename):
		# Make a temporary file, open it for reading and writing, and 
		# dump the decoded attachment payload into it.
		temp = tempfile.NamedTemporaryFile()
		temp_path = temp.name
		temp.write(part.get_payload(decode=True))
		
		# Open temp file with olefile, get the metadata, and test the
		# comments value, which is what we will put the RID into.
		# This requires a placeholder value that we test for because
		# OLE files are very fragile and it must be replaced with a 
		# new value of the same length as the placeholder.
		ole = olefile.OleFileIO(temp_path, write_mode=True)
		meta = ole.get_metadata()
		comments = meta.comments
		comments_len = len(comments)
		
		if comments == placeholder:
			# This stream has the file properties we want
			data = ole.openstream('\x05SummaryInformation').read()
			# Replace with new value, padding it to the same length as
			# the placeholder with filler characters (padding variable).
			data_new = data.replace(comments, bytearray(rid.rjust(comments_len, padding), 'utf-8'))
			ole.write_stream('\x05SummaryInformation', data_new)
			ole.close()
		else:
			quit(2) # Placeholder value not found; whoops
		
		temp.seek(0) # Return read head to start before updating attachment
		part.set_content(temp.read(), maintype='application', subtype='octet-stream', filename=filename)
		temp.close()
		
		# Dump message back to stdout for exim
		print(msg.as_string())
