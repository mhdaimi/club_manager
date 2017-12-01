from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle, TA_CENTER
from reportlab.lib.units import inch, mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Table, SimpleDocTemplate, Spacer
from pyPdf import PdfFileWriter, PdfFileReader
import datetime
 
########################################################################
class Test(object):
    """"""
 
    #----------------------------------------------------------------------
    def __init__(self):
        """Constructor"""
        self.width, self.height = letter
        self.styles = getSampleStyleSheet()
 
    #----------------------------------------------------------------------
    def coord(self, x, y, unit=1):
        """
        http://stackoverflow.com/questions/4726011/wrap-text-in-a-table-reportlab
        Helper class to help position flowables in Canvas objects
        """
        x, y = x * unit, self.height -  y * unit
        return x, y
 
    #----------------------------------------------------------------------
    def run(self):
        """
        Run the report
        """
        self.doc = SimpleDocTemplate("test.pdf")
        self.story = [Spacer(1, 1*inch)]
        self.createLineItems()
 
        self.doc.build(self.story, onFirstPage=self.first_page, onLaterPages=self.later_page)
        print "finished!"
        
        with open("test.pdf", "rb") as f:
            print "merginig"
            new_pdf = PdfFileReader(f)
            existing_pdf = PdfFileReader(file("report_template.pdf", "rb"))
            output = PdfFileWriter()
            page = existing_pdf.getPage(0)
            page.mergePage(new_pdf.getPage(0))
            output.addPage(page)
            
            for page in range(new_pdf.getNumPages()-1):
                output.addPage(new_pdf.getPage(page+1))
            
            outputStream = file("final_report.pdf", "wb")
            output.write(outputStream)
            outputStream.close()
 
    #----------------------------------------------------------------------
    def later_page(self, canvas, doc):
        canvas.setFont("Times-Roman", 10)
        page_num = canvas.getPageNumber()
        text = "Page #%s" % page_num
        canvas.drawRightString(200*mm, 20*mm, text)
        
    def first_page(self, canvas, doc):
        
        canvas.setFont("Times-Roman", 10)
        canvas.drawString(65,720,"Finance report for : %s %s" %(datetime.date.today().strftime("%B"), datetime.date.today().strftime("%Y")))
        canvas.drawString(65,705,'Report generated on : %s' %str("{:%d.%m.%Y}".format(datetime.datetime.now())))
        canvas.setLineWidth(.2)
        canvas.line(50,695,540,695)
        
        page_num = canvas.getPageNumber()
        text = "Page #%s" % page_num
        canvas.drawRightString(200*mm, 20*mm, text)
 
    #----------------------------------------------------------------------
    def createLineItems(self):
        """
        Create the line items
        """
        inp_data = [("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),
        ("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),
        ("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),
        ("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),
        ("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),
        ("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),
        ("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),
        ("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),
        ("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),
        ("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),
        ("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20")]
        text_data = ["Sr. No", "Ref. No", "Name", "Expected Amt.", "Paid Amt."]
        d = []
        font_size = 11
        centered = ParagraphStyle(name="centered", alignment=TA_CENTER)
        for text in text_data:
            ptext = "<font name=Times-Roman size=%s><b>%s</b></font>" % (font_size, text)
            p = Paragraph(ptext, centered)
            d.append(p)
 
        data = [d]
 
        line_num = 1
 
        formatted_line_data = []
 
        for each_rec in inp_data:
            line_data = [str(line_num),each_rec[0],each_rec[1],each_rec[2],each_rec[3]]
 
            for item in line_data:
                ptext = "<font name=Times-Roman size=%s>%s</font>" % (font_size-1, item)
                p = Paragraph(ptext, centered)
                formatted_line_data.append(p)
            data.append(formatted_line_data)
            formatted_line_data = []
            line_num += 1
 
        table = Table(data, colWidths=[75, 75, 150, 100, 100])
 
        self.story.append(table)
 
#----------------------------------------------------------------------
if __name__ == "__main__":
    t = Test()
    t.run()