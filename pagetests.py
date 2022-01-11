from docx import Document

document = Document()


name = input('Name: ')
protocol = input('Protocol: ')

paragraph = document.add_paragraph(f'Python is cool!!\n'
    f'This document was created by {name}, with the protocol {protocol}')

document.save(f'documents/document_{name}-{protocol}.docx')