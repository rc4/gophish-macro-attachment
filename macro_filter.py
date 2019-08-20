#!/usr/bin/python3.6

import sys, re, email, tempfile
import olefile
from email.parser import BytesParser
from email import policy

padding = b'.'
placeholder = b'{{.RidPlaceholder}}'

EXT_RE = re.compile(r'(\.doc|\.xls)$', re.IGNORECASE)
RID_RE = re.compile(r'<!--\040RID:\040([a-z0-9]{6,})\040-->', re.IGNORECASE)
MHT_RE = re.compile(r'\.mht$', re.IGNORECASE)
MHT_PLACEHOLDER_RE = re.compile(b'{{\.RIDPLACEHOLDER}}')

msg = email.message_from_file(sys.stdin, policy=policy.default)

for part in msg.walk():
	filename = part.get_filename() # Get filename of part
	content_main = part.get_content_maintype()
	content_sub = part.get_content_subtype()

	if not filename and content_main != 'multipart': # Body will not have a filename
		content = part.get_content()
		matches = RID_RE.search(content) # Search for RID string in content
		if matches:
			rid_bytes = bytes(matches.group(1), 'utf-8')
			new_content = RID_RE.sub('', content)
			part.set_content(new_content, subtype=content_sub)
		else:
			continue
	elif not filename:
		continue
	elif EXT_RE.search(filename):
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
			data_new = data.replace(comments, rid_bytes.rjust(comments_len, padding))
			ole.write_stream('\x05SummaryInformation', data_new)
			ole.close()
		
		temp.seek(0) # Return read head to start before updating attachment
		part.set_content(temp.read(), maintype=content_main, subtype=content_sub, filename=filename)
		temp.close()
	elif MHT_RE.search(filename):
		mht_filename = MHT_RE.sub('.doc', filename) # Replace .mht with .doc
		mht_content = part.get_payload(decode=True)
		new_mht = MHT_PLACEHOLDER_RE.sub(rid_bytes, mht_content)
		part.set_content(new_mht, maintype=content_main, subtype=content_sub, cte="base64", filename=mht_filename)
		
# Dump message back to stdout for exim
print(msg.as_string())
