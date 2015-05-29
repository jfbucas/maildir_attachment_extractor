#!/usr/bin/env python3

"""Extract MIME attachments from a Maildir file into a folder structure YYYY/MM/DD/"""

import os
import sys
import email
import errno
import mimetypes
import base64
import unicodedata
import codecs
import datetime
from zipfile import ZipFile

from argparse import ArgumentParser

# Global variables - not so elegant, but works at this scale
lines = None
boundaries =  None
nb_boundaries = 0 
username = None
mailfolder = None
filename = ""
filebase64 = False

# -------------
# Merge the line counter and the lines
def merge_o( o, ct ):
	o[ 0 ] += ct[ 0 ]
	o[ 1 ].extend( ct[ 1 ] )
	return o

# -------------
def get_line( nb_l ):
	global lines
	if nb_l < len(lines):
		return  lines[nb_l]
	else:
		return ""

# -------------
def get_next_line( nb_l ):

	ct = get_content_type( nb_l + 1 )
	if ct != None:
		return ct

	ct = get_content_name( nb_l + 1 )
	if ct != None:
		return ct

	ct = get_content_description( nb_l + 1 )
	if ct != None:
		return ct

	ct = get_content_encoding( nb_l + 1 )
	if ct != None:
		return ct

	ct = get_content_disposition( nb_l + 1 )
	if ct != None:
		return ct

	ct = get_x_attachment( nb_l + 1 )
	if ct != None:
		return ct

	ct = get_rt_attachment( nb_l + 1 )
	if ct != None:
		return ct

	ct = get_skip_line( nb_l + 1 )
	if ct != None:
		return ct

	ct = get_external_attachment( nb_l + 1 )
	if ct != None:
		return None

	return None

# -------------
def get_boundary( nb_l ):
	global boundaries
	global nb_boundaries

	l=get_line( nb_l )
	for boundary in boundaries:
		if boundary in l:
			#print('\t+ Boundary #' + str(nb_boundaries) + ' found')
			nb_boundaries += 1

			o = [ 1, [ l ] ]

			ct = get_next_line( nb_l )
			if ct != None:
				return merge_o(o, ct)
	return None

# -------------
def get_content_type( nb_l ):
	l=get_line( nb_l )
	if l.lower().startswith('content-type: '):
		if l.lower().startswith('content-type: text/plain'):
			print('\t> Content-Type: text, skipping')
			return None
		elif l.lower().startswith('content-type: application/pgp-signature'):
			print('\t> Content-Type: PGP, skipping')
			return None
		#elif l.lower().startswith('content-type: application/x-zip-compressed'):
		#	print('Content-Type: Zipped, skipping')
		#elif l.lower().startswith('content-type: audio/mpg'):
		#	print('Content-Type: MPG, skipping')
		#elif l.lower().startswith('content-type: image/gif'):
		#	print('Content-Type: GIF, skipping')
		#elif l.lower().startswith('content-type: image/png'):
		#	print('Content-Type: PNG, skipping')

		#o = [ 1, [ "Content-Type: application/x-zip-compressed;\n" ] ]
		print("\t>", l, end="")
		o = [ 1, [ l ] ]

		ct = get_next_line( nb_l )
		if ct != None:
			return merge_o(o, ct)

		ct = get_content_type_options( nb_l + 1 )
		if ct != None:
			return merge_o(o, ct)

	return None

# -------------
def get_content_type_options( nb_l ):
	global filename
	l=get_line( nb_l )

	if l.strip().lower().startswith('x-mac-creator='):
		pass
	elif l.strip().lower().startswith('x-mac-type='):
		pass
	elif l.strip().lower().startswith('x-mac-hide-extension='):
		pass
	elif l.strip().lower().startswith('x-unix-mode='):
		pass
	elif l.strip().lower().startswith('name='):
		pass
	else:
		return None

	o = [ 1, [ l ] ]

	ct = get_next_line( nb_l )
	if ct != None:
		return merge_o(o, ct)

	ct = get_content_type_options( nb_l + 1 )
	if ct != None:
		return merge_o(o, ct)

	return None


# -------------
def get_content_name( nb_l ):
	l=get_line( nb_l )
	if l.strip().lower().startswith('name="'):
		#o = [ 1, [ l[:-2]+'.zip"\n' ] ]
		o = [ 1, [ l ] ]

		ct = get_next_line( nb_l )
		if ct != None:
			return merge_o(o, ct)
	return None


# -------------
def get_content_encoding( nb_l ):
	global filebase64
	l=get_line( nb_l )
	if l.lower().startswith('content-transfer-encoding: base64'):
		print("\t>", l, end="")
		filebase64 = True
		o = [ 1, [ l ] ]

		ct = get_next_line( nb_l )
		if ct != None:
			filebase64 = False
			return merge_o(o, ct)

		filebase64 = False

	return None

# -------------
def get_content_disposition( nb_l ):
	global filename
	l=get_line( nb_l )
	if l.lower().startswith('content-disposition:'):
		o = [ 1, [ l ] ]
		print("\t>", l, end="")

		if 'filename=' in l:
			filename = l.split('filename=')[1].strip().replace('\"','')

		ct = get_next_line( nb_l )
		if ct != None:
			return merge_o(o, ct)
			
		ct = get_content_disposition_options( nb_l + 1 )
		if ct != None:
			return merge_o(o, ct)

	
	return None

# -------------
def get_content_disposition_options( nb_l ):
	global filename
	l=get_line( nb_l )

	if l.strip().lower().startswith('filename="'):
		#o = [ 1, [ l[:-2]+'.zip"\n' ] ]
		filename = l.split('"')[1]
		print("\t> Filename: "+filename)
	elif l.strip().lower().startswith('filename*0="'):
		#o = [ 1, [ l[:-2]+'.zip"\n' ] ]
		filename = l.split('"')[1]
		print("\t> Filename: "+filename)
	elif l.strip().lower().startswith('filename*1="'):
		#o = [ 1, [ l[:-2]+'.zip"\n' ] ]
		filename += l.split('"')[1]
		print("\t> Filename: "+filename)
	elif l.strip().lower().startswith('filename='):
		#o = [ 1, [ l[:-2]+'.zip"\n' ] ]
		filename = l.split('=')[1].strip()
		print("\t> Filename: "+filename)
	elif l.strip().lower().startswith('creation-date="'):
		#print("\t>", l, end="")
		pass
	elif l.strip().lower().startswith('modification-date="'):
		#print("\t>", l, end="")
		pass
	elif l.strip().lower().startswith('size='):
		#print("\t>", l, end="")
		pass
	else:
		return None

	o = [ 1, [ l ] ]

	ct = get_next_line( nb_l )
	if ct != None:
		return merge_o(o, ct)

	ct = get_content_disposition_options( nb_l + 1 )
	if ct != None:
		return merge_o(o, ct)

	return None

# -------------
def get_content_description( nb_l ):
	l=get_line( nb_l )
	if l.lower().startswith('content-description:'):
		print("\t>", l, end="")
		o = [ 1, [ l ] ]

		ct = get_next_line( nb_l )
		if ct != None:
			return merge_o(o, ct)

		
	return None


# -------------
def get_external_attachment( nb_l ):

	l=get_line( nb_l )
	if l.strip().lower().startswith('x-mozilla-external-attachment-url:'):
		o = [ 1, [ l ] ]
		print("\t> Already extracted, skipping")
		return o

	return None

# -------------
def get_x_attachment( nb_l ):

	l=get_line( nb_l )
	if l.strip().lower().startswith('x-attachment-id:'):
		o = [ 1, [ l ] ]

		ct = get_next_line( nb_l )
		if ct != None:
			return merge_o(o, ct)

	return None

# -------------
def get_rt_attachment( nb_l ):

	l=get_line( nb_l )
	if l.strip().lower().startswith('rt-attachment:'):
		o = [ 1, [ l ] ]

		ct = get_next_line( nb_l )
		if ct != None:
			return merge_o(o, ct)

	return None

# -------------
def get_skip_line( nb_l ):
	l=get_line( nb_l )
	if l == "\n":
		o = [ 1, [ l ] ]

		ct = get_content_uu( nb_l + 1 )
		if ct != None:

			#ct = get_packed( ct )
			#if ct != None:
			#	return merge_o(o, ct)
			ct = get_detach( ct )
			if ct != None:
				# if detach is successful, the empty line is included in ct
				#return merge_o(o, ct)
				return ct

		#return o
		
	return None

# -------------
def get_content_uu( nb_l ):
	global boundaries
	l=get_line( nb_l )
	n = 0
	uu_l = []
	while (l != "\n") and (l != ""):
		for boundary in boundaries:
			if boundary in l:
				if len(uu_l) > 0:
					return [ n, uu_l ]

				return None
				

		uu_l.append( l )

		nb_l += 1
		l=get_line( nb_l )

		n += 1

	if len(uu_l) > 0:
		return [ n, uu_l ]

	return None

# -------------
def get_packed( ct ):
	global filename

	# Write attachment
	s = ""
	for l in ct[1]:
		s += l.strip()
	dec = base64.standard_b64decode(s)
	
	with open( "zipped/" + filename, "bw" ) as fp:
		fp.write( dec )

	# Zip attachment
	with ZipFile( 'zipped/' + filename + '.zip', 'w') as myzip:
	    myzip.write('zipped/' + filename )

	# Base64 Attachment
	enc = ""
	with open( "zipped/" + filename + '.zip', "br" ) as fp:
		enc = base64.standard_b64encode( fp.read() ).decode('utf-8')

	print( '   >> size(Unzipped - Unzipped) = ', os.path.getsize( "zipped/" + filename ) - os.path.getsize( "zipped/" + filename + '.zip') )
	if os.path.getsize( "zipped/" + filename + '.zip') < os.path.getsize( "zipped/" + filename ):
		os.remove( "zipped/" + filename )
		# split the base64 of the new file into len=72 strings
		ct1 = []
		w = 72
		for n in range(len(enc)//w):
			ct1.append( enc[n*w:(n+1)*w-1] + "\n" )

		if (len(enc) % w) != 0:
			n = (len(enc)//w) 
			ct1.append( enc[n*w:] + "\n" )
		
		return [ ct[0], ct1 ]
	else:
		os.remove( "zipped/" + filename + '.zip' )

	return None

# -------------
def get_detach( ct ):
	global mailfolder, username, filename, filebase64

	if filename == "":
		return None

	if not os.path.exists( mailfolder):
		os.makedirs( mailfolder )

	# Write attachment
	s = ""
	dec = ""
	if filebase64:
		for l in ct[1]:
			s += l.strip()
		dec = base64.standard_b64decode(s)
		with open( mailfolder + "/" + filename, "bw" ) as fp:
			fp.write( dec )
	else:
		for l in ct[1]:
			s += l
		with open( mailfolder + "/" + filename, "w" ) as fp:
			fp.write( s )



	ct[1] = [
		"X-Mozilla-External-Attachment-URL: file:///home/" + username + "/" + mailfolder+"/"+filename+"\n\n",
		"The attachment was detached from this message and placed in the folder /home/" + username + "/" + mailfolder+"/"+filename+"\n", 
	]

	return ct

# -------------
def main():
	global lines, boundaries, mailfolder, username

	parser = ArgumentParser(description="""\
		Extract attachment for a MIME message.
		""")

	parser.add_argument('msgfile')
	args = parser.parse_args()

	print( " o Working on "+ args.msgfile )
	
	try:
		username = args.msgfile.split("/")[0]
		maildate = datetime.datetime.fromtimestamp( int(args.msgfile.split("/")[-1].split(".")[0]) )
		mailfolder = "attachments/"+ str(maildate.year)+"/"+str(maildate.month).rjust(2,'0')+"/"+str(maildate.day).rjust(2,'0')
	except ValueError:
		print( "############[ "+ args.msgfile +" incorrect name ]======================" )
		#print( username, maildate, mailfolder ) 
		return
		

	encoding=""
	try:
		encoding="utf-8"
		lines = open(args.msgfile, 'r', encoding=encoding).readlines()
	except UnicodeDecodeError:
		try:
			encoding="iso-8859-1"
			lines = open(args.msgfile, 'r', encoding=encoding).readlines()
		except UnicodeDecodeError:
			print( "############[ "+ args.msgfile +" ]======================" )


	if lines == None:
		print( "@@@@@@@@@[ "+ args.msgfile +" empty ]" )
		return

	# Look for boundary markers
	nb_l=0
	boundaries = []
	while nb_l < len(lines):
		l=get_line( nb_l )
		if l.lower().startswith('content-type: multipart/mixed; boundary="'):
			boundaries.append( l.split('"')[1] )
			print("\t- Boundary is ", boundaries[-1])
		elif l.lower().startswith('content-type: multipart/mixed; boundary='):
			boundaries.append( l.split('=')[1].strip() )
			print("\t- Boundary is ", boundaries[-1])
		elif l.lower().startswith('content-type: multipart/mixed;'):
			l2=get_line( nb_l+1 )
			if l2.strip().startswith('boundary="'):
				boundaries.append( l2.split('"')[1] )
				print("\t- Boundary is ", boundaries[ -1 ])
			elif l2.strip().startswith('boundary='):
				boundaries.append( l2.split('=')[1].strip() )
				print("\t- Boundary is ", boundaries[ -1 ])

		nb_l+=1

	# Get the work done on the multipart
	if len(boundaries) > 0:

		nb_l = 0
		repack = 0
		output = []
		while nb_l < len(lines):

			o = get_boundary( nb_l )
			if o != None:
				nb_l += o[ 0 ]
				output.extend( o[ 1 ] )
				repack += 1
			else:
				output.append( lines[nb_l] )
				nb_l += 1

		if repack > 0:
			#print( "  > Repacked !" )
			print( "\t>>> Extracted " + str(repack) + " !" )
			#for l in output:
			#	print( l, end="" )
			#os.rename( args.msgfile, args.msgfile + ".old" )
			os.remove( args.msgfile )
			with open( args.msgfile, "w", encoding=encoding ) as fp:
				for l in output:
					fp.write( l )



if __name__ == '__main__':
	main()
