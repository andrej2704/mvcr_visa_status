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
          if "soubor" in link['href']:
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
def sendmail():
    gmail_user = "yourusername@gmail.com"
    gmail_pwd = "password"
    TO = 'recipient@gmail.com'
    SUBJECT = "MVCR:Your request was approved!"
    TEXT = "Yahoo!! Dance, open champagne!"
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login(gmail_user, gmail_pwd)
    BODY = '\r\n'.join(['To: %s' % TO,
                        'From: %s' % gmail_user,
                        'Subject: %s' % SUBJECT,
                        '', TEXT])

    server.sendmail(gmail_user, [TO], BODY)
    print ('email sent')


def searchForID(id, page, filename=XLS_FIL_ENAME):
  book = open_workbook(filename)
  sheet = book.sheets()[int(page)]
  print u"Excel sheet name: {}".format(sheet.name)
  found = False

  for r in range(sheet.nrows):
      for c in range(sheet.ncols):
          cell = sheet.cell(r, c)
          if id in cell.value:
              found = True
              print u"Your number was found!!!, here it is: {}".format(sheet.row(r)[1].value)
              sendmail()
              break

  if not found: print "Your number was not found :("

def main(id, page):
  url = getDownloadUrlFromMvcr()
  downloadXlsFileFromMvcr(url)
  searchForID(id, page)

if __name__ == "__main__":
  parser = argparse.ArgumentParser(description="MVCR visa application results")
  parser.add_argument("-id", "--id", default="OAM-27/TP-2017", help="Provide your registered number")
  parser.add_argument("-page", "--page", default="0", help="Provide excel page number. 0 - Long Term Residence (Default), 1 - Work Permit, 2 - Permanent Residence")

  args = parser.parse_args()
  main(args.id, args.page)
