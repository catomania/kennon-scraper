import requests, os, bs4, openpyxl
from openpyxl import Workbook

# if no workbook exists, make a new one
if os.path.isfile("kennon_links.xlsx") is True:
	pass
else:	
	# make an workbook and add column headers
	wb = Workbook()

	dest_filename = "kennon_links.xlsx"

	ws1 = wb.active
	ws1.title = "links and description"

	wb.save(filename = dest_filename)

row = 2 # record the first row to start inputting data

# go through all the pages and grab the links
for x in range(1, 208):

	url = "http://www.joshuakennon.com/page/" + str(x) + "/" #starting URL, loops through to the last page on the site
	print "grabbing " + str(url)
	

	# get the URL
	res = requests.get(url)
	res.raise_for_status()

	soup = bs4.BeautifulSoup(res.text)

	#find article links
	linkElem = soup.find_all("a", href=True, attrs={"rel": "bookmark"})
	if linkElem == []: # check if any elements match our search terms
		print "not found!"
	else:
		wb = openpyxl.load_workbook('kennon_links.xlsx') # open up the wb and start adding data
		sheet = wb.get_active_sheet()
		sheet['A1'] = "Link"	
		
		for l in linkElem:

			if '<a class="data-link' in str(l): # skip the data-link redundant links
				pass
			else: 
				cell = "A" + str(row) # add in a href links to our wb
				sheet[cell] = str(l)
				row += 1
				print row
		wb.save('kennon_links.xlsx') # save the wb when you're all done!









