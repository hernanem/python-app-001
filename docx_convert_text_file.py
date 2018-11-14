import os
import win32com.client

#needed the full file path not just the file name even in the same directory
DOC_FILEPATH = "C:\Users\Edwin\Documents\Python\AI Test Plan GBS_AISIT7_Final.docx"
doc = win32com.client.GetObject(DOC_FILEPATH)
text = doc.Range().Text

if 'TABLE OF CONTENTS' in text:
    print 'Table of Contents = True'
else:
    print 'Table of Contents = False'

if 'OVERVIEW' in text:
    print 'OVERVIEW = True'
else:
    print 'OVERVIEW = False'

if 'PURPOSE' in text:
    print 'PURPOSE = True'
else:
    print 'PURPOSE = False'

if 'OBJECTIVE' in text:
    print 'OBJECTIVE = True'
else:
    print 'OBJECTIVE = False'

if 'SCOPE' in text:
    print 'SCOPE = True'
else:
    print 'SCOPE = False'

if 'TEST SCHEDULE' in text:
    print 'TEST SCHEDULE = True'
else:
    print 'TEST SCHEDULE = False'

if 'TEST PERSONNEL' in text:
    print 'TEST PERSONNEL = True'
else:
    print 'TEST PERSONNEL = False'

if 'REFERENCE DOCUMENTS' in text:
    print 'REFERENCE DOCUMENTS = True'
else:
    print 'REFERENCE DOCUMENTS = False'

if 'RISKS' in text:
    print 'RISKS = True'
else:
    print 'RISKS = False'

if 'CONCEPT OF OPERATIONS' in text:
    print 'CONCEPT OF OPERATIONS = True'
else:
    print 'CONCEPT OF OPERATIONS = False'

if 'SYSTEM DESCRIPTION' in text:
    print 'SYSTEM DESCRIPTION = True'
else:
    print 'SYSTEM DESCRIPTION = False'

if 'KEY FEATURES AND SUBSYSTEMS' in text:
    print 'KEY FEATURES AND SUBSYSTEMS = True'
else:
    print 'KEY FEATURES AND SUBSYSTEMS = False'

if 'PHYSICAL INTERFACES' in text:
    print 'PHYSICAL INTERFACES = True'
else:
    print 'PHYSICAL INTERFACES = False'

if 'LOGICAL INTERFACES' in text:
    print 'LOGICAL INTERFACES = True'
else:
    print 'LOGICAL INTERFACES = False'

if 'TEST OVERVIEW' in text:
    print 'TEST OVERVIEW = True'
else:
    print 'TEST OVERVIEW = False'

if 'PREVIOUS TEST EVENTS' in text:
    print 'PREVIOUS TEST EVENTS = True'
else:
    print 'PREVIOUS TEST EVENTS = False'



with open("plain_text.txt", "wb") as f:
    f.write(text.encode("utf-8"))

os.startfile("plain_text.txt")
