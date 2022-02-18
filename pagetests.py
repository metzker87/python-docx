from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor, Inches

document = Document()
sections = document.sections
for section in sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches (1)
    section.right_margin = Inches(1)

name = input('Name: ')
protocol = input('Protocol: ')

paragraph = document.add_paragraph(f'Python is cool!!\n'
    f'This document was created by {name}, with the protocol {protocol}')

paragraph.style = document.styles.add_style('normalText', WD_STYLE_TYPE.PARAGRAPH)
font = paragraph.style.font
font.name = 'Arial'
font.size = Pt(11)
#font.bold = True
#font.italic = True
#font.underline = True
font.color.rgb = RGBColor(0, 0, 0)

paragraph_format = paragraph.paragraph_format
paragraph_format.space_before = Pt(12)
paragraph_format.space_after = Pt(0)


document.save(f'documents/document_{name}-{protocol}.docx')