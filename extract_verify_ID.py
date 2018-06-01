## Excel worksheet ID extraction script
## By Bingyin Hu 05/25/2018

import xlrd
import sys

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
        message += '[Sample ID Error] Sample ID format error: Sample ID has extra parts, should be of format "L101_S1_LastName_2018". Current upload is "%s".\n' % (ID_raw)
    elif len(ID_seg) < 4:
        message += '[Sample ID Error] Sample ID format error: Sample ID has missing parts, should be of format "L101_S1_LastName_2018". Current upload is "%s".\n' % (ID_raw)
    else:
        # PID
        PID = ID_seg[0]
        if PID[0].isalpha():
            # PID starts with the wrong alphabet
            if PID[0] not in ['L','E']:
                message += '[PID Error] Sample ID format error: PID must start with "L" (for literature data) or "E" (for experimental data) case-sensitive. Current upload starts with "%s". Example: "L101".\n' % (PID[0])
            # PID length
            if len(PID) < 4:
                message += '[PID Error] Sample ID format error: PID must have at least a length of 4. Current upload has a length of "%s". Example: "L101".\n' % (len(PID))
            # PID ends with non-digits
            elif not PID[1:].isdigit():
                message += '[PID Error] Sample ID format error: PID must end with numbers. Current upload ends with "%s". Example: "L101".\n' % (PID[1:])
        else:
            # PID starts with non-alphabet
            message += '[PID Error] Sample ID format error: PID must start with "L" (for literature data) or "E" (for experimental data). Current upload is missing the alphabet. Example: "L101".\n'
        # SID
        SID = ID_seg[1]
        if SID[0].isalpha():
            # SID starts with the wrong alphabet
            if SID[0] != 'S':
                message += '[SID Error] Sample ID format error: SID must start with "S" case-sensitive. Current upload starts with "%s". Example: "S7".\n' % (SID[0])
            # SID length
            if len(SID) < 2:
                message += '[SID Error] Sample ID format error: SID must have at least a length of 2. Current upload has a length of "%s". Example: "S7".\n' % (len(SID))
            # SID ends with non-digits
            elif not SID[1:].isdigit():
                message += '[SID Error] Sample ID format error: SID must end with numbers. Current upload ends with "%s". Example: "S7".\n' % (SID[1:])
        else:
            # SID starts with non-alphabet
            message += '[SID Error] Sample ID format error: SID must start with "S". Current upload is missing the alphabet. Example: "S7".\n'
        # AuthorLastName
        ALN = ID_seg[2]
        # PubYear
        PubYear = ID_seg[3]
        if not PubYear.isdigit():
            message += '[PubYear Error] Sample ID format error: Publication year must be a year. Current upload is "%s". Example: "2018".\n' % (PubYear)
    return message
        
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
        message += '[Excel Error] Excel template format error: Sample_Info sheet not found.\n'
        with open('./ID.txt', 'w') as fid:
            fid.write(message)
        return
    # otherwise, find and save the ID in ./ID.txt
    for row in xrange(sheet_sample.nrows):
        # ID
        if match(sheet_sample.row_values(row)[0], 'Sample ID'):
            ID_raw = str(sheet_sample.row_values(row)[1])
            # if no ID is entered in the cell
            if len(ID_raw.strip()) == 0:
                message += '[Excel Error] Excel template value error: Sample ID is not entered in the uploaded Excel template.\n'
            # else verify the entered ID
            else:
                message += verifyID(ID_raw)
    # if no error detected
    if message == '':
        # write the ID in ./ID.txt
        with open('./ID.txt', 'w') as fid:
            fid.write(ID_raw)
    else:
        # write the message in ./error_message.txt
        with open('./error_message.txt', 'a') as fid:
            fid.write(message)
    return
 
xlsxName = './'+sys.argv[1] # sys.argv[1] command line action
extractID(xlsxName)
