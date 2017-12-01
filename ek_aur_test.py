# import time
# from reportlab.lib.enums import TA_JUSTIFY, TA_RIGHT
# from reportlab.lib.pagesizes import letter
# from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
# from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
# from reportlab.lib.units import inch
# from reportlab.pdfgen import canvas
# 
# #!/usr/bin/env python
# # -*- coding: utf-8 -*-
# 
# lyrics = "This is some sort of string. I dont know how this will work, let us see what happens now! I have now got the address right. need to see how other things go!"
# 
# canvas = canvas.Canvas("test.pdf", pagesize=letter, bottomup=1)
# canvas.setFont("Times-Roman", 8)
# 
# 
# canvas.drawString(75,705,'Pak Dar Ul-Islam Islamische-Gemeinde e.V.')
# canvas.drawString(75,695,'Muenchener Str. 55, 60329 Frankfurt am Main')
# canvas.setLineWidth(.3)
# canvas.line(75,690,250,690)
# canvas.setFont("Times-Roman", 11)
# canvas.drawString(75,670,'Mohammad Daimi')
# canvas.drawString(75,655,'Winterstr 19,')
# canvas.drawString(75,640,'60489 Frankfurt')
# canvas.setFont("Times-Roman", 10)
# canvas.drawString(75,550,lyrics)
# 
# Story=[]
# formatted_time = time.ctime()
# styles=getSampleStyleSheet()
# styles.add(ParagraphStyle(name='Justify', alignment=TA_JUSTIFY))
# ptext = '<font size=12>%s</font>' % formatted_time
#  
# Story.append(Paragraph(ptext, styles["Normal"]))
# Story.append(Spacer(1, 12))
#     
# canvas.save()

# from lxml import etree
# from docx import Document
# document = Document()
# r = 2 # Number of rows you want
# c = 2 # Number of collumns you want
# table = document.add_table(rows=r, cols=c)
# cell = table.cell(0, 1)
# table.style = 'LightShading-Accent1' # set your style, look at the help documentation for more help
# for y in range(r):
#     for x in range(c):
#         cell.text = 'text goes here'
# document.save('demo.docx') # Save document
# 
#!/usr/bin/env python
# -*- coding: utf-8 -*-
from pyPdf import PdfFileWriter, PdfFileReader
import StringIO, copy
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
import win32api
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.rl_config import defaultPageSize
from reportlab.lib.units import inch, cm
from reportlab.lib.pagesizes import letter
from reportlab.lib.enums import TA_JUSTIFY

# name_wali_pdf = PdfFileReader(file("temp_doc.pdf", "rb+"))
# 
# existing_pdf = PdfFileReader(file("masjid_template.pdf", "rb+"))
# output = PdfFileWriter()
# # add the "watermark" (which is the new pdf) on the existing page
# page = existing_pdf.getPage(0)
# 
# 
# page.mergePage(name_wali_pdf.getPage(0))
# output.addPage(page)
# # finally, write "output" to a real file
# outputStream = file("final.pdf", "wb")
# output.write(outputStream)
# outputStream.close()
# #win32api.ShellExecute(0, "print", "final.pdf", None,  ".",  0)

text = """This is a line. This is a line. This is a line. This is a line. This is a line. This is a line. This is a line. This is a line. This is a line. 
This is a line. This is a line. This is a line. This is a line. This is a line. This is a line. This is a line. This is a line. This is a line. 
This is a line. This is a line. This is a line. This is a line. This is a line. This is a line.This is a line. This is a line.This is a line. 
This is a line. This is a line. This is a line. This is a line. This is a line. This is a line. This is a line. This is a line.
"""
packet = StringIO.StringIO()
# create a new PDF with Reportlab
width, height = A4
can = canvas.Canvas(packet, pagesize=A4)

can.setFont("Times-Roman", 11)
can.drawString(55,670,"Mohammad Daimi")
can.drawString(55,655,"Winterstr 19,")
can.drawString(55,640,"60489 Frankfurt am Main")

styles = getSampleStyleSheet()
        
#doc = SimpleDocTemplate("temp_doc.pdf", pagesize=letter)
Story = [Spacer(1,3*inch)]
style = styles["Normal"]
p = Paragraph(text, style)
Story.append(p)

print width, height

p.wrap(width-100, 500)

p.drawOn(can, width-540, 500)

#can.drawString(75,550,Story)

can.save()

#move to the beginning of the StringIO buffer
packet.seek(0)
new_pdf = PdfFileReader(packet)
# read your existing PDF
existing_pdf = PdfFileReader(file("masjid_template.pdf", "rb"))
output = PdfFileWriter()
# add the "watermark" (which is the new pdf) on the existing page
page = existing_pdf.getPage(0)
page.mergePage(new_pdf.getPage(0))
output.addPage(page)
# finally, write "output" to a real file
outputStream = file("destination.pdf", "wb")
output.write(outputStream)
outputStream.close()

