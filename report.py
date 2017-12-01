from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.rl_config import defaultPageSize
from reportlab.lib.units import inch
PAGE_HEIGHT=defaultPageSize[1]; PAGE_WIDTH=defaultPageSize[0]
styles = getSampleStyleSheet()
from pyPdf import PdfFileWriter, PdfFileReader
bogustext="""
I have changed my working PC few months ago, and as you probably know your self, it takes time to install all software that you need. Usually you do that when you need it. So this days there was need for me to make some pdf files, and haven't had acrobat pro installed yet. Usually all other softwares I use have option to export to pdf, but I need to merge all files in one file, and for this purpose acrobat pro is convenient option. Since only thing needed to be done is merging pdf files in one, I thought maybe write some small script in python. After maybe half on hour of searching for appropriate python lib for this task, I had my python pdf script with under ten lines of code. 
I have changed my working PC few months ago, and as you probably know your self, it takes time to install all software that you need. Usually you do that when you need it. So this days there was need for me to make some pdf files, and haven't had acrobat pro installed yet. Usually all other softwares I use have option to export to pdf, but I need to merge all files in one file, and for this purpose acrobat pro is convenient option. Since only thing needed to be done is merging pdf files in one, I thought maybe write some small script in python. After maybe half on hour of searching for appropriate python lib for this task, I had my python pdf script with under ten lines of code.
"""

def myFirstPage(canvas, doc):
    canvas.saveState()
#     canvas.setFont('Times-Bold',16)
#     canvas.drawCentredString(PAGE_WIDTH/2.0, PAGE_HEIGHT-108, Title)
#     canvas.setFont('Times-Roman',9)
#     canvas.drawString(inch, 0.75 * inch, "First Page / %s" % pageinfo)
#     canvas.setFont("Times-Roman", 8)
#     canvas.drawString(75,705,'Pak Dar Ul-Islam Islamische-Gemeinde e.V.')
#     canvas.drawString(75,695,'Muenchener Str. 55, 60329 Frankfurt am Main')
#     canvas.setLineWidth(.3)
#     canvas.line(75,690,250,690)
    canvas.setFont("Times-Roman", 10)
    canvas.drawString(55,670,'Mohammad Daimi')
    canvas.drawString(55,655,'Winterstr 19,')
    canvas.drawString(55,640,'60489 Frankfurt')
    canvas.drawString(55,590,'Dear Mohammad Daimi,')
    
#     canvas.setFont("Times-Roman", 10)
#     canvas.drawString(350,750,'Pak Dar Ul-Islam Moschee Frankfurt')
#     canvas.setFont("Times-Roman", 10)
#     canvas.drawString(350,735,'Muenchener Str. 55, 60329 Frankfurt am Main')
#     canvas.drawString(350,720,'Email: pakdarulislam@gmail.com')
#     canvas.drawString(350,705,'Tel: 069/47869781 - 84')
#     canvas.drawString(350,690,'Tel: 069/235688')
#     canvas.drawString(350,675,'Fax: +49 (069) 232062')
    
#     canvas.setFont('Times-Roman',10)
#     canvas.drawString(inch, 0.80 * inch, "Bankverbindung: National Bank")   
#     canvas.drawString(inch, 0.65 * inch, "Konto Nr.: 0081904005")
#     canvas.drawString(inch, 0.50 * inch, "BLZ: 50130000")
#     canvas.drawString(inch, 0.35 * inch, "IBAN: DE67501300000081904005")
#     canvas.drawString(inch, 0.20 * inch, "BIC: NBPADEFF")
    canvas.restoreState()
    
def go():
    doc = SimpleDocTemplate("new.pdf")
    Story = [Spacer(0,2.7*inch)]
    style = styles["Normal"]
    
    ptext = '<font name=Times-Roman size=10>%s</font>' % bogustext
    
    p = Paragraph(ptext, style)
    Story.append(p)
    Story.append(Spacer(1,0.5*inch))
    doc.build(Story, onFirstPage=myFirstPage)
    
    
    new_pdf = PdfFileReader(file("new.pdf", "rb"))
    existing_pdf = PdfFileReader(file("masjid_template.pdf", "rb"))
    output = PdfFileWriter()
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    outputStream = file("ek_nayi_File.pdf", "wb")
    output.write(outputStream)
    outputStream.close()

go()