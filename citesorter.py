import sys
import xml.etree.ElementTree as ET
import uuid
import re
from docx import Document

# Portions gratefully appropriated from https://github.com/JStrydhorst/pubmed-to-word 
# TODO - Pare down cloned code to just necessary parsing

def parse_nbib(lines):
  # reads the lines of text until it finds a tag in the first four columns, then
  # concatenates the values until the start of the next tag
  nlines = 0
  tag = ''
  for line in lines:
    if len(tag) == 0:
      tag = line[0:4].strip()
      val = line[5:].strip()
      nlines += 1
    elif not line[0:4].strip():
      val = val + ' ' + line[5:].strip()
      nlines += 1
    else:
      break    
  return tag, val, nlines

def import_sources(filename):
  # converts the citation file to MSWord compatible bibliography xml file
  sources = ET.Element('b:Sources',
                       {"SelectedStyle" : "",
                        "xmlns:b" : "http://schemas.openxmlformats.org/officeDocument/2006/bibliography",
                        "xmlns" : "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"})
  f = open(filename, 'rU')
  lines = f.readlines()
  line = 0
  while line < len(lines):
    tag, val, linesread = parse_nbib(lines[line:])
    line += linesread
    if tag == 'PMID':
      source = ET.SubElement(sources,'b:Source')
      sourcetag = ET.SubElement(source, 'b:Tag')
      sourcetype = ET.SubElement(source, 'b:SourceType')
      sourcetype.text = 'JournalArticle'
      sourceguid = ET.SubElement(source, 'b:Guid')
      sourceguid.text = '{' + str(uuid.uuid4()) + '}'
      sourcetitle = ET.SubElement(source, 'b:Title')
      sourceyear = ET.SubElement(source, 'b:Year')
      sourcejournal = ET.SubElement(source, 'b:JournalName')
      sourcepages = ET.SubElement(source, 'b:Pages')
      sourceauthor = ET.SubElement(source, 'b:Author')
      sourceauthor2 = ET.SubElement(sourceauthor, 'b:Author')
      sourcenames = ET.SubElement(sourceauthor2, 'b:NameList')
      sourcevolume = ET.SubElement(source, 'b:Volume')
    elif tag == 'TI':
      sourcetitle.text = val
    elif tag == 'TA':
      sourcejournal.text = val
    elif tag == 'DP':
      sourceyear.text = val[0:4]
    elif tag == 'FAU':
      match = re.search(r'^([\w\'-]+),\s*([\w\'-]+)\s*(.+)?' ,val)
      if match:
        person = ET.SubElement(sourcenames, 'b:Person')
        lastname = ET.SubElement(person, 'b:Last')
        lastname.text = match.group(1)
        if match.group(3):
          middlename = ET.SubElement(person, 'b:Middle')
          middlename.text = match.group(3)
        firstname = ET.SubElement(person, 'b:First')
        firstname.text = match.group(2)
        if not sourcetag.text:
          sourcetag.text = match.group(1)[0:3] + sourceyear.text[2:4]
    elif tag == 'VI':
      sourcevolume.text = val
    elif tag == 'PG':
      sourcepages.text = val
  f.close()
  tree = ET.ElementTree(sources)
  tree.write('sources.xml',encoding="utf-8",xml_declaration=True)

# EXTRACT DATA
import_sources("citations.nbib")

# FORMAT DATA

# EXPORT DATA
document = Document('CA_Template.docx')
# add info to template here
#TODO - Look into exporting under individual headings in doc rather than having to recreate doc completely

document.save('output.docx') #TODO - Automate this to have the current date as the filename




