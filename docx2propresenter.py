file = 'acts2billv.docx'
#file = 'casst.docx'
#file = 'test3.docx'
# file = 'OVERHEAD.docx'
proFile = 'pythonTest.pro5'

import references as ref
from docx import Document
import pro5

# standardConversion
# ------------------------------
# Word has random characters, this gets
# rid of some of those so that python doesn't
# freak out so much and we don't have weird characters
# returns the original text with standard characters replaced
def standardConversion(text):
	text = text.replace( u'[]', u'')
	text = text.replace( u'\u2013', u'-')
	text = text.replace( u'\u2014', u'-')
	text = text.replace( u'\u2018', u"'")
	text = text.replace( u'\u2019', u"'")
 	text = text.replace( u'\u201c', u'"')
 	text = text.replace( u'\u201d', u'"')
	text = text.replace( u'\u2026', u'...')
	text = text.replace( u'\xa0', u' ') # nbsp

	return text

# initialise variables
# html = ''
htmlPara = '' # the formatted slide
slides = [] # an array of slides
currentSlide = ['', ''] # a slide.. [0] is formatted, [1] is text

document = Document(file)
for para in document.paragraphs:
	htmlPara = ''
	for run in para.runs:
		start = ''
		end = ''

		if run.bold:
			start += '\\f1\\b'
			end += '\\f0\\b0'
		if run.italic:
			start += '\\i'
			end += '\\i0'

		htmlPara += start + ' ' + run.text + end
		currentSlide[1] += run.text
	htmlPara += '\\\n'
	currentSlide[0] += htmlPara

	if htmlPara == '\\\n': # empty paragraph
		if (currentSlide[0][:-4] != '' and currentSlide[0][:-4] != ' '):
			slides.append([currentSlide[0][:-4], currentSlide[1]]) # get rid of the last new line
		currentSlide = ['',''] # reset current slide


index = 0
f = open(proFile, 'wa')
f.write(pro5.headers())

for slide in slides:
	slide[0] = standardConversion(slide[0])
	slide[1] = standardConversion(slide[1])
	scriptures = ref.extract(slide[0]) # a tuple containing scriptures found
	
	f.write(pro5.textslide(str(index), slide[1][0:10], slide[0]))
	index += 1

f.write(pro5.closure())
f.close()
