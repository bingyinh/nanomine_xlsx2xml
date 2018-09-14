# NanoMine xlsx to xml conversion tool

By Bingyin Hu

### 1. System preparations

Required packages:

- glob
  - Python default package

- os
  - Python default package

- sys
  - Python default package

- copy
  - Python default package

- time
  - Python default package

- datetime
  - Python default package

- xlrd
  - https://github.com/python-excel/xlrd
  - Read the input Excel files.

- dicttoxml
  - https://pypi.org/project/dicttoxml/
  - Convert python dictionary to xml.

- collections
  - Python default package

- xml.etree.ElementTree
  - Python default package

- pickle
  - Python default package

- csv
  - Python default package

- lxml
  - Used in the validation module.
  - http://lxml.de/

- Beautiful Soup 4
  - https://www.crummy.com/software/BeautifulSoup/bs4/doc/index.html
  - Used in the DOI modules.
    
- mechanize
  - https://github.com/sparklemotion/mechanize
  - "Mechanize is a ruby library that makes automated web interaction easy."
  - Used in the DOI modules.

- ast
  - Python default package
  - Used in the DOI modules.

Open the command or terminal and run
```
pip install -r requirements.txt
```
### 2. How to run

1. Add the downloaded directory to the sys.path. Note that the NanoMine xml schema is not provided in this repository. It can be downloaded at https://github.com/Duke-MatSci/nanomine-schema/tree/master/xml.

2. Apply for an account at https://apps.crossref.org/requestaccount/ for the Crossref Query Services and save the email address in `downloaded_directory/account.txt` as required by the DOI query module.

2. Assign values to:
   - `jobDir`: the directory of the Excel files and other files that are to be converted.
   - `code_srcDir`: the directory of the downloaded codes (current directory).
   - `xsdDir`: the directory of the xml schema to be validated against.
   - `templateName`: the file name of the Excel template.

3. In python, run the `extract_verify_ID_callable.py` by
```
from extract_verify_ID_callable import runEVI
runEVI(jobDir, code_srcDir, xsdDir, templateName)
```

4. If there is no `error_message.txt` generated in the `jobDir` and an `ID.txt` is generated, the conversion can be kicked off by
```
from customized_compiler_callable import compiler
logName = compiler(jobDir, code_srcDir, xsdDir, templateName)
```
where `logName` is the directory for the schema validation error log.

5. Check the error log for potential schema validation error. There should be an `/xml` folder generated in the `jobDir`, and the converted xml file will be inside.
