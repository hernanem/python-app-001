#import elementtree.ElementTree as ET
#import elementtree
#import xml.etree.ElementTree as ET

def check():
    #tree = ET.parse("test.xml")
    #doc = tree.getroot()
    #thingy = doc.find('timeSeries')

    datafile = file('AI Test Plan GBS_AISIT7_Final.docx')
    found = 'False'
    for line in datafile:
        if 'APPLICATION' in line:
            found = 'True'
            break
    print found
    #print thingy.attrib
    #print doc[0][1].text


#print thingy.attrib
#print doc[0][1].text

check()
