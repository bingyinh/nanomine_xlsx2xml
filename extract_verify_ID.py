## Excel worksheet ID extraction script
## By Bingyin Hu 05/25/2018

# -------------------------------------------- file os and other general lib
import glob, os                           # Python standard library
import sys                                # Python standard library
import copy                               # Python standard library
from time import gmtime, strftime         # Python standard library

# -------------------------------------------- convert .XLSX to .XML files
import xlrd
import dicttoxml
import uuid
import codecs

# -------------------------------------------- XML Schemas
from xml.dom.minidom import parseString   # Python standard library
import collections                        # Python standard library
import xml.etree.ElementTree as ET

# -------------------------------------------- DOI information retriever
from doiretriever import mainDOI          # extract publication info by DOI

# -------------------------------------------- XML-XSD validation tool
from xml_update_validator import runValidation


# a helper method to find a blurry match regardless of # signs between two
# strings, testant is the standard expression
def match(testee, testant):
    if (testant.lower() == testee.lower()):
        return True
    elif (testant.lower() == testee.lower().split("#")[0].strip()):
        return True
    return False

# the method to verify ID
def verifyID(ID_raw):
    message = '' # init
    ID_seg = ID_raw.split('_') # a standard ID format PID_SID_AuthorLastName_PubYear
    if len(ID_seg) > 4:
        message = '[Error02] Sample ID has extra parts, should be of format "L101_S1_LastName_2018". Current upload is "%s".' % (ID_raw)
    elif len(ID_seg) < 4:
        message = '[Error03] Sample ID has missing parts, should be of format "L101_S1_LastName_2018". Current upload is "%s".' % (ID_raw)
    else:
        PID = ID_seg[0]
        if PID[0].isalpha():
            if PID[0].lower() not in ['l','e']:
                message = '[Error04] Sample ID format error: PID must start with "L" (for literature data) or "E" (for experimental data). Current upload starts with "%s". Example: "L101".\n' % (PID[0])
            



# the method to extract ID
def extractID(xlsxName):
    # open xlsx
    # xlrd is the library used to read xlsx file
    # https://secure.simplistix.co.uk/svn/xlrd/trunk/xlrd/doc/xlrd.html?p=4966
    xlfile = xlrd.open_workbook(xlsxName)
    # find the sheet with ID
    sheet_sample = '' # init
    sheets = xlfile.sheets()
    for sheet in sheets:
        # check the header of the sheet to determine what it has inside
        if (sheet.row_values(0)[0].strip().lower() == "sample info"):
            sheet_sample = sheet
    # if the sheet with ID is not found, write error message in ./ID.txt
    message = ''
    if sheet_sample == '':
        message = '[Error00] Sample_Info sheet not found'
        with open('./ID.txt', 'w') as fid:
            fid.write(message)
        return
    # otherwise, find and save the ID in ./ID.txt
    for row in xrange(sheet_sample.nrows):
        # ID
        if match(sheet.row_values(row)[0], 'Sample ID'):
            ID_raw = str(sheet.row_values(row)[1])
            # if no ID is entered in the cell
            if len(ID_raw).strip() == 0:
                message = '[Error01] Sample ID is not entered in the uploaded Excel template'
            # else verify the entered ID
            else:
                message = verifyID(ID_raw)
    # write the message in ./ID.txt    
    with open('./ID.txt', 'w') as fid:
        fid.write(message)
    return
 
xlsxName = './'+sys.argv[1] # sys.argv[1] command line action
extract(xlsxName)