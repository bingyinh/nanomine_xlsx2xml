## Excel worksheet ID extraction script
## By Bingyin Hu 05/25/2018

import xlrd
import sys
from doiretriever import mainDOIsoupFirst
import pickle
import xml.etree.ElementTree as ET
import dicttoxml
import collections

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
    # ID_seg = ID_raw.split('_') # a standard ID format PID_SID_AuthorLastName_PubYear
    # if len(ID_seg) > 4:
    #     message += '[Sample ID Error] Sample ID format error: Sample ID has extra parts, should be of format "L101_S1_LastName_2018". Current upload is "%s".\n' % (ID_raw)
    # elif len(ID_seg) < 4:
    #     message += '[Sample ID Error] Sample ID format error: Sample ID has missing parts, should be of format "L101_S1_LastName_2018". Current upload is "%s".\n' % (ID_raw)
    # else:
    #     # PID
    #     PID = ID_seg[0]
    #     if PID[0].isalpha():
    #         # PID starts with the wrong alphabet
    #         if PID[0] not in ['L','E']:
    #             message += '[PID Error] Sample ID format error: PID must start with "L" (for literature data) or "E" (for experimental data) case-sensitive. Current upload starts with "%s". Example: "L101".\n' % (PID[0])
    #         # PID length
    #         if len(PID) < 4:
    #             message += '[PID Error] Sample ID format error: PID must have at least a length of 4. Current upload has a length of "%s". Example: "L101".\n' % (len(PID))
    #         # PID ends with non-digits
    #         elif not PID[1:].isdigit():
    #             message += '[PID Error] Sample ID format error: PID must end with numbers. Current upload ends with "%s". Example: "L101".\n' % (PID[1:])
    #     else:
    #         # PID starts with non-alphabet
    #         message += '[PID Error] Sample ID format error: PID must start with "L" (for literature data) or "E" (for experimental data). Current upload is missing the alphabet. Example: "L101".\n'
    # # SID
    # SID = ID_seg[1]
    SID = ID_raw
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
    # # AuthorLastName
    # ALN = ID_seg[2]
    # # PubYear
    # PubYear = ID_seg[3]
    # if not PubYear.isdigit():
    #     message += '[PubYear Error] Sample ID format error: Publication year must be a year. Current upload is "%s". Example: "2018".\n' % (PubYear)
    return message
        
# the method to extract ID
def extractID(xlsxName, myXSDtree):
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
        # DOI
        if match(sheet_sample.row_values(row)[0], 'DOI'):
            DOI = str(sheet_sample.row_values(row)[1])
    # if no error detected
    if message == '':
        # call localDOI here
        localdoiDict = localDOI(DOI, myXSDtree)
        # generate ID here
        newID = generateID(localdoiDict, ID_raw)
        # write the ID in ./ID.txt
        with open('./ID.txt', 'w') as fid:
            fid.write(newID)
    else:
        # write the message in ./error_message.txt
        with open('./error_message.txt', 'a') as fid:
            fid.write(message)
    return


# check local dict for doi info
def localDOI(DOI, myXSDtree):
    with open('doi.pkl','rb') as f:
        alldoiDict = pickle.load(f)
    if DOI not in alldoiDict:
        # assign it 'nextPID', update 'nextPID', save it into alldoiDict, update
        # doi.pkl, fetching the metadata is slow, so we need to make sure the
        # paperID is updated in the doi.pkl first to avoid collision.
        PID = alldoiDict['nextPID']
        alldoiDict['nextPID'] += 1
        alldoiDict[DOI] = {'paperID':PID}
        with open('doi.pkl', 'wb') as f:
            pickle.dump(alldoiDict, f)
        # now fetch the metadata using doi-crawler and save to alldoiDict, doi.pkl
        crawlerDict = mainDOIsoupFirst(DOI)
        # transfer the newdoiDict to an xml element
        tree = dict2element(crawlerDict, myXSDtree) # an xml element
        citation = tree.find('.//Citation')
        alldoiDict[DOI]['metadata'] = citation
        # update the doi.pkl for the metadata field
        with open('doi.pkl', 'wb') as f:
            pickle.dump(alldoiDict, f)
        return alldoiDict[DOI]
    else:
        return alldoiDict[DOI]

# generate ID with format PID_SID_LastName_PubYear for users with DOI
def generateID(doiDict, SID):
    PID = doiDict['paperID']
    LastName = 'LastName'
    Name = doiDict['metadata'].find('.//Author')
    if Name is not None:
        LastName = Name.text.split(',')[0]
    PubYear = 'PubYear'
    PubYearRaw = doiDict['metadata'].find('.//PublicationYear')
    if PubYearRaw is not None:
        PubYear = PubYearRaw.text
    return '_'.join([str(PID), SID, LastName, PubYear])

# convert DOI crawler dict into an xml element
def dict2element(crawlerDict, myXSDtree):
    # init
    CommonFields = []
    Journal = []
    Citation = collections.OrderedDict()
    CitationType = collections.OrderedDict()
    output = collections.OrderedDict()
    # port dict infos into lists
    for key in crawlerDict:
        if key == "ISSN" or key == "Issue":
            if len(crawlerDict[key]) > 0:
                Journal.append({key: crawlerDict[key][0]})
        elif key == "Author" or key == "Keyword":
            if len(crawlerDict[key]) > 0:
                for value in crawlerDict[key]:
                    CommonFields.append({key: value})
        elif key == "Institution":
            if len(crawlerDict[key]) > 0:
                CommonFields.append({u"Location": crawlerDict[key][0]})
        else:
            if len(crawlerDict[key]) > 0:
                CommonFields.append({key: crawlerDict[key][0]})
    # sort sequence
    CommonFields = sortSequence(CommonFields, 'CommonFields', myXSDtree)
    Journal = sortSequence(Journal, 'Journal', myXSDtree)
    # save to a dict
    if len(CommonFields) > 0:
        Citation[u'CommonFields'] = CommonFields
    if len(Journal) > 0:
        CitationType = collections.OrderedDict([(u'Journal',Journal)])
    if len(CitationType) > 0:
        Citation[u'CitationType'] = CitationType
    if len(Citation) > 0:
        output = collections.OrderedDict([(u'Citation', Citation)])
    # convert to an xml element
    assert (len(output) > 0)
    doi_xml = dicttoxml.dicttoxml(output,attr_type=False)
    doi_xml = doi_xml.replace('<item>','').replace('</item>','')
    tree = ET.ElementTree(ET.fromstring(doi_xml))
    return tree

################################################################################
# copied from customize_compiler_ubuntu.py
# a helper method to sort a list of dicts according to their keys, the sequence
# is extracted from the xsd file
def sortSequence(myList, myClassName, myXSDtree):
    # don't sort if input is not a list
    if type(myList) != list:
        return myList
    CTmap = {'Distribution':'Distribution',# small_list_data;'Distribution'
             'CommonFields':'CitationDetailType',# CommonFields;'CommonFields'
             'Journal':'CitationJournalType',# Journal;'Journal'
             'LabGenerated':'LabGeneratedType',# LabGenerated;'LabGenerated'
             'MATERIALS':'MaterialsType', # temp_list;'MATERIALS'
             'Matrix':'MatrixType',# temp;prevTemp
             'MatrixComponent':'PolymerType',# MatrixComponent;'MatrixComponent'
             'Filler':'FillerType',# temp;prevTemp
             'FillerComponent':'ParticleType',# FillerComponent;'FillerComponent'
             'Particle Surface Treatment (PST)':'ParticleSurfaceModificationType',# PST;'Particle Surface Treatment (PST)'
             'PROCESSING':'ProcessingCategoryType',# Process_list;'PROCESSING'
             'Additive':'ProcessingParameterAdditive',# temp;prevTemp
             'Solvent':'SolventType',# temp;prevTemp
             'Mixing':'ProcessingParameterMixing',# temp;prevTemp
             'Extrusion':'ExtrusionType',# temp;prevTemp not necessary
             'SingleScrewExtrusion':'SingleScrewExtrusionExtruderType',# Extrusion;prevExtrsHeader
             'TwinScrewExtrusion':'TwinScrewExtrusionExtruderType',# Extrusion;prevExtrsHeader
             'ExtrusionHeatingZone':'HeatingZoneType',# ExtrsHZ;'ExtrusionHeatingZone'
             'ExtrusionOutput':'ExtruderOutput',# ExtrsOP;'ExtrusionOutput'
             'Heating':'GeneralConditionsType',# temp;prevTemp
             'Cooling':'GeneralConditionsType',# temp;prevTemp
             'Drying/Evaporation':'GeneralConditionsType',# temp;prevTemp
             'Molding':'MoldingDescriptionType',# temp;prevTemp
             'MoldingInfo':'GeneralConditionsType',# MoldingInfo;'MoldingInfo'
             'CHARACTERIZATION':'MeasurementMethodsType',# temp_list;'CHARACTERIZATION'
             'Transmission electron microscopy':'SEMTEMtype',# temp;prevTemp
             'Scanning electron microscopy':'SEMTEMtype',# temp;prevTemp
             'Dynamic mechanical analysis':'GeneralEquipmentType',# temp;prevTemp
             'Dielectric and impedance spectroscopy analysis':'GeneralEquipmentType',# temp;prevTemp
             'Differential scanning calorimetry':'DSCMeasurementType',# temp;prevTemp
             'Fourier transform infrared spectroscopy':'GeneralEquipmentType',# temp;prevTemp
             'Xray diffraction and scattering':'GeneralEquipmentType',# temp;prevTemp
             'Xray photoelectron spectroscopy':'GeneralEquipmentType',# temp;prevTemp
             'Atomic force microscopy':'GeneralEquipmentType',# temp;prevTemp
             'Thermogravimetric analysis':'GeneralEquipmentType',# temp;prevTemp
             'Raman spectroscopy':'GeneralEquipmentType',# temp;prevTemp
             'Nuclear magnetic resonance':'GeneralEquipmentType',# temp;prevTemp
             'Pulsed electro acoustic':'GeneralEquipmentType',# temp;prevTemp
             'Rheometry':'RheometryType',# temp;prevTemp
             'Electrometry':'GeneralEquipmentType',# temp;prevTemp
             'Optical microscopy':'GeneralEquipmentType',# temp;prevTemp
             'PROPERTIES':'PropertiesType',# DATA_PROP;'PROPERTIES'
             'Mechanical':'MechanicalType',# temp_list;'Mechanical'
             'Tensile':'TensileType',# temp;prevTemp
             'Conditions':'TensileMeasurementMethodType',# Conditions;'Conditions'
             'Flexural':'FlexuralType',# temp;prevTemp
             'Compression':'CompressionType',# temp;prevTemp
             'Shear':'ShearType',# temp;prevTemp
             'Fracture':'FractureToughnessType',# temp;prevTemp
             'Essential work of fracture (EWF)':'EWFType',# tempFracture;prevTempFrac
             'Linear Elastic':'LinearElasticType',# tempFracture;prevTempFrac
             'Plastic Elastic':'PlasticElasticType',# tempFracture;prevTempFrac
             'Impact':'ImpactEnergyType',# temp;prevTemp
             'Viscoelastic':'ViscoelasticType',# temp_list;'Viscoelastic'
             'Dynamic properties':'DMAType',# temp;prevTemp
             'Frequency sweep':'DMAFrequencyType',# DMA_Test;prevDMA
             'Temperature sweep':'DMATemperatureType',# DMA_Test;prevDMA
             'Strain sweep':'DMAStrainSweepTypeType',# DMA_Test;prevDMA
             'Creep':'CreepType',# temp;prevTemp
             'CompressiveVisc':'CompressiveCreepType',# temp_Creep;prevCreep
             'TensileVisc':'TensileCreepType',# temp_Creep;prevCreep
             'FlexuralVisc':'FlexuralCreepType',# temp_Creep;prevCreep
             'Electrical':'DielectricPropertiesType', # temp_list;'Electrical'
             'AC dielectric dispersion':'DielectricDispersionType', # temp;prevTemp
             'Dielectric breakdown strength':'BreakdownStrengthType', # temp;prevTemp
             'Thermal':'ThermalType', # temp_list;'Thermal'
             'Crystallinity':'CrystallizationType', # temp;prevTemp
             'Volumetric':'VolumetricType', # temp_list;'Volumetric'
             'Rheological':'RheologicalType', # temp_list;'Rheological'
             'MICROSTRUCTURE':'MicrostructureDescriptorsType', # temp_list;'MICROSTRUCTURE'
             'Imagefile':'ImageFileType', # temp;prevTemp
             'Dimension':'ImageSizeType', # Dimension;'Dimension'
             'Sample experimental info':'SampleType', # temp;prevTemp
             'PolymerNanocomposite':'Root' # DATA;'PolymerNanocomposite'
             } # {myClassName:xsdComplexTypeName}
    myTypeName = CTmap[myClassName] #example: MatrixType
    myTypeTree = myXSDtree.findall(".//*[@name='" + myTypeName + "']")
    # prepare seq dict {className:index}
    index = 0
    seq = {}
    for ele in myTypeTree[0].iter('{http://www.w3.org/2001/XMLSchema}element'):
        seq[ele.get('name')] = index
        index += 1
    # sort myList by the index for each element key
    myList.sort(key=lambda x: seq[x.keys()[0]])
    return myList
################################################################################

if __name__ == '__main__':
    # read the xsd tree
    xsdDir = "./PNC_schema_060718.xsd"
    # xsdDir = './'+sys.argv[2]
    myXSDtree = ET.parse(xsdDir)
    xlsxName = './'+sys.argv[1] # sys.argv[1] command line action
    extractID(xlsxName, myXSDtree)