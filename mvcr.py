import requests
import argparse
from xlrd import open_workbook
from BeautifulSoup import BeautifulSoup, SoupStrainer

XLS_FIL_ENAME = "prehled.xls"

def getDownloadUrlFromMvcr():
  response = requests.get('http://www.mvcr.cz/clanek/informace-o-stavu-rizeni.aspx')
  url = ""
  for link in BeautifulSoup(response.content, parseOnlyThese=SoupStrainer('a')):
      if link.has_key(u'href'):
          if "prehled" in link['href']:
              url = "http://www.mvcr.cz/" + link['href']
              print "Download url: {}".format(url)
              return url

def downloadXlsFileFromMvcr(url, filename=XLS_FIL_ENAME):

  response = requests.get(url)

  with open(filename, 'wb') as f:
      for chunk in response.iter_content(chunk_size=1024):
          if chunk:
              f.write(chunk)
      f.close()

  return filename

def searchForID(id, filename=XLS_FIL_ENAME):
  book = open_workbook(filename)
  sheet = book.sheets()[2]
  print u"Excel sheet name: {}".format(sheet.name)
  found = False

  for r in range(sheet.nrows):
      for c in range(sheet.ncols):
          cell = sheet.cell(r, c)
          if id in cell.value:
              found = True
              print u"Your number was found!!!, here it is: {}".format(sheet.row(r)[1].value)
              break

  if not found: print "Your number was not found :("

def main(id):
  url = getDownloadUrlFromMvcr()
  downloadXlsFileFromMvcr(url)
  searchForID(id)

if __name__ == "__main__":
  parser = argparse.ArgumentParser(description="MVCR visa application results")
  parser.add_argument("-id", "--id", default="import requests
import argparse
from xlrd import open_workbook
from BeautifulSoup import BeautifulSoup, SoupStrainer

XLS_FIL_ENAME = "prehled.xls"

def getDownloadUrlFromMvcr():
  response = requests.get('http://www.mvcr.cz/clanek/informace-o-stavu-rizeni.aspx')
  url = ""
  for link in BeautifulSoup(response.content, parseOnlyThese=SoupStrainer('a')):
      if link.has_key(u'href'):
          if "prehled" in link['href']:
              url = "http://www.mvcr.cz/" + link['href']
              print "Download url: {}".format(url)
              return url

def downloadXlsFileFromMvcr(url, filename=XLS_FIL_ENAME):

  response = requests.get(url)

  with open(filename, 'wb') as f:
      for chunk in response.iter_content(chunk_size=1024):
          if chunk:
              f.write(chunk)
      f.close()

  return filename

def searchForID(id, filename=XLS_FIL_ENAME):
  book = open_workbook(filename)
  sheet = book.sheets()[2]
  print u"Excel sheet name: {}".format(sheet.name)
  found = False

  for r in range(sheet.nrows):
      for c in range(sheet.ncols):
          cell = sheet.cell(r, c)
          if id in cell.value:
              found = True
              print u"Your number was found!!!, here it is: {}".format(sheet.row(r)[1].value)
              break

  if not found: print "Your number was not found :("

def main(id):
  url = getDownloadUrlFromMvcr()
  downloadXlsFileFromMvcr(url)
  searchForID(id)

if __name__ == "__main__":
  parser = argparse.ArgumentParser(description="MVCR visa application results")
  parser.add_argument("-id", "--id", default="OAM-3295/TP-2017", help="Provide your registered number")
  args = parser.parse_args()
  main(args.id)







