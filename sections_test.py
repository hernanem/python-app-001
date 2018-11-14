import docx

document = docx.Document('AI Test Plan GBS_AISIT7_Final.docx')
sections = document.sections

print len(sections)
