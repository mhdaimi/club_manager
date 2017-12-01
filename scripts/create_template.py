#!/usr/bin/python
# -*- coding: utf-8 -*-
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.rl_config import defaultPageSize
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import letter, A4
from pyPdf import PdfFileWriter, PdfFileReader
import StringIO
from reportlab.pdfgen import canvas


canvas = canvas.Canvas("pdf_template.pdf", pagesize=A4)
canvas.setFont("Times-Roman", 8)
canvas.drawString(55,705,'ABCD HOLDINGS PVT LTD')
canvas.drawString(55,695, "xyz street 55, 60000 Berlin")
canvas.setLineWidth(.2)
canvas.line(55,690,250,690)
#canvas.setFont("Times-Roman", 11)
#canvas.drawString(75,670,self.p_name)
#canvas.drawString(75,655,self.p_street)
#canvas.drawString(75,640,self.p_city)

#canvas.drawImage("new_logo_masjid.jpg", 55, 750, 490, 40)
canvas.drawString(255, 750,"Your company Logo here")

canvas.setFont("Times-Roman", 10)
canvas.drawString(350,705,'ABCD HOLDINGS PVT LTD')
canvas.setFont("Times-Roman", 10)
canvas.drawString(350,690,'xyz street 55, 60000 Berlin')
canvas.drawString(350,675,'Web: www.abcdholdings.com')
canvas.drawString(350,660,'Email: xyz@abcdholdings.com')
canvas.drawString(350,645,'Tel: 00/0123456789, 01/0123456789')
#canvas.drawString(350,630,'Tel: 01/0123456789')
canvas.drawString(350,630,'Fax: +00 0123456789')


canvas.setLineWidth(.2)
canvas.line(55,1.3*inch,540,1.3*inch)
canvas.setFont('Times-Roman',10)
canvas.drawString(inch, 1 * inch, "Bankverbindung: Bank name")   
canvas.drawString(inch, 0.80 * inch, "Konto Nr.: Account number here")
canvas.drawString(inch, 0.60 * inch, "BLZ: Bank BLZ")
canvas.drawString(inch, 0.40 * inch, "IBAN: BANK IBAN")
canvas.drawString(inch, 0.20 * inch, "BIC: BANK BIC")

# canvas.drawString(315, 1 * inch, "Vorstand: Präsident Mohammad Azam Shaheen")
# canvas.drawString(315, 0.80 * inch, "SteuerNr: 4725007020")
# canvas.drawString(315, 0.60 * inch, "Amtsgericht Frankfurt am Main Vereinsregister Nr: 9140")   


canvas.save()