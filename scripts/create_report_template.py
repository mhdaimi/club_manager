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


canvas = canvas.Canvas("report_template.pdf", pagesize=A4)
# canvas = canvas.Canvas("masjid_template.pdf", pagesize=A4)
# canvas.setFont("Times-Roman", 8)
# canvas.drawString(55,705,'Pak Dar Ul-Islam Islamische-Gemeinde e.V.')
# canvas.drawString(55,695, "M�nchener Str. 55, 60329 Frankfurt am Main")
# canvas.setLineWidth(.2)
# canvas.line(55,690,250,690)
#canvas.setFont("Times-Roman", 11)
#canvas.drawString(75,670,self.p_name)
#canvas.drawString(75,655,self.p_street)
#canvas.drawString(75,640,self.p_city)

canvas.drawString(255, 750, "YOUR COMPANY NAME & LOGO")

# canvas.setFont("Times-Roman", 10)
# canvas.drawString(350,705,'Pak Dar Ul-Islam Islamische-Gemeinde e.V.')
# canvas.setFont("Times-Roman", 10)
# canvas.drawString(350,690,'M�nchener Str. 55, 60329 Frankfurt am Main')
# canvas.drawString(350,675,'Web: http://pakdarulislam.com/')
# canvas.drawString(350,660,'Email: pakdarulislam@gmail.com')
# canvas.drawString(350,645,'Tel: 069/47869781 - 84')
# canvas.drawString(350,630,'Tel: 069/235688')
# canvas.drawString(350,615,'Fax: +49 (069) 232062')
# 
# 
# canvas.setLineWidth(.2)
# canvas.line(55,1.3*inch,540,1.3*inch)
# canvas.setFont('Times-Roman',10)
# canvas.drawString(inch, 1 * inch, "Bankverbindung: National Bank")   
# canvas.drawString(inch, 0.80 * inch, "Konto Nr.: 0081904005")
# canvas.drawString(inch, 0.60 * inch, "BLZ: 50130000")
# canvas.drawString(inch, 0.40 * inch, "IBAN: DE67501300000081904005")
# canvas.drawString(inch, 0.20 * inch, "BIC: NBPADEFF")
# 
# canvas.drawString(315, 1 * inch, "Vorstand: Pr�sident Mohammad Azam Shaheen")
# canvas.drawString(315, 0.80 * inch, "SteuerNr: 4725007020")
# canvas.drawString(315, 0.60 * inch, "Amtsgericht Frankfurt am Main Vereinsregister Nr: 9140")   


canvas.save()