#!/usr/bin/python

import os

dirlist = ['SWPPP','NOI','sitemap','NOC','CSN','CSN2018R','CSNDI','MS4S2018R','MS4Submittal','NOI2018R','NOT','Permit','Permit2018R','PermitCity','PondCalculations','SoilData','Delegation Letter']

cwd = os.getcwd()
fid = open('corrupt_pdf_list.txt', 'w')
fid.write("Finding files for directory {0}\n".format(cwd))
tcnt = 0
ccnt = 0
for dir in dirlist:
	curdir = os.path.join(cwd,'images',dir)
	print('=============\n{0}\n=============\n'.format(curdir))
	fid.write('=============\n{0}\n=============\n'.format(curdir))
	for file in os.listdir(curdir):
		#print(file)
		tcnt = tcnt + 1
		if file.endswith(".pdf"):
			with open(os.path.join(curdir,file), 'rb') as f:
				f.seek(-2, os.SEEK_END)
				while f.read(1) != b'\n':
					f.seek(-2, os.SEEK_CUR)
				try:
					last_line = f.readline().decode()
					eof_str = last_line.strip()[-5:]
					if eof_str != "%%EOF":
						ccnt = ccnt + 1
						print('{0} - {1} is bad: {2}\n'.format(ccnt,file,eof_str))
						fid.write('{0} of {1} - {2} is bad: {3}\n'.format(ccnt,tcnt,file,eof_str))
				except:
					print('{0} failed\n'.format(file))
					fid.write('{0} failed\n'.format(file))
fid.close()