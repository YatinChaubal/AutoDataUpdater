import urllib2
import datetime
import zipfile
import os
from win32com.client import Dispatch



class NseBhavcopy:
	def __init__(self,Date=datetime.date.today()):
	
		curdt=Date
		yy=curdt.year
		mm=curdt.strftime("%b")
		dd=curdt.day
		

		if dd < 10:
			dd ="0"+str(dd)
		url="http://www.nseindia.com/content/historical/EQUITIES/"+str(yy)+"/"+str(mm).upper()+"/cm"+str(dd)+str(mm).upper()+str(yy)+"bhav.csv.zip"
		
	
		hdr = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11',
			'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
			'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3',
			'Accept-Encoding': 'none',
			'Accept-Language': 'en-US,en;q=0.8',
			'Connection': 'keep-alive'}
		req = urllib2.Request(url, headers=hdr)

		fileavailable=1
		try:
			page = urllib2.urlopen(req)
		except urllib2.HTTPError, e:
			#print e.fp.read()
			fileavailable=0
			self.filename=""
			self.ValidData=0
			

		if fileavailable == 1:
			meta = page.info()
			file_size = int(meta.getheaders("Content-Length")[0])
			file_name = url.split('/')[-1]
			f = open(file_name, 'wb')
			file_size_dl = 0
			block_sz = 8192
			while True:
				buffer = page.read(block_sz)
				if not buffer:
					break

				file_size_dl += len(buffer)
				f.write(buffer)
				status = r"%10d  [%3.2f%%]" % (file_size_dl, file_size_dl * 100. / file_size)
				status = status + chr(8)*(len(status)+1)
			f.close()
		
			zfile=open(file_name, 'rb')
			z=zipfile.ZipFile(zfile)
			for name in z.namelist():
				outfile = open(name, 'wb')
				outfile.write(z.read(name))
				outfile.close()
			zfile.close()
			self.filename=name
			self.ValidData=1
		
			

		
		
def CSVToAmibroker(fname,amidb,format_file):

	ab = Dispatch("Broker.Application")
	ab.LoadDatabase(amidb)
	ab.Import(0, fname, format_file)
	ab.SaveDatabase()
	print ab
	print "Done"

	
CSVToAmibroker("C:\\NSE Auto Downloader\\cm15APR2014bhav.csv","C:\\Amibroker DB\\NSE testdb","NSE.format")


	
