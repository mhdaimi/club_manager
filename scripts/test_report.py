import os,sys
import Tkinter as tk
import tkFileDialog, tkMessageBox
import sqlite3, string, re, datetime
from pyPdf import PdfFileWriter, PdfFileReader
from reportlab.platypus import Paragraph, Table, SimpleDocTemplate, Spacer
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle, TA_CENTER
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch,mm
import arabic_reshaper
from bidi.algorithm import get_display
from docx import Document, text
from docx.shared import Pt
from docx.enum.text import *


data = [("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),("0001", "Mohammad Daimi", "25", "20"),("0002", "Mohammad Daimi", "25", "20"),
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








def create_report():
    """
    Run the report
    """
    self.doc = SimpleDocTemplate("test.pdf")
    self.story = [Spacer(1, 1*inch)]
    self.createLineItems()

    self.doc.build(self.story, onFirstPage=self.add_pagenum, onLaterPages=self.add_pagenum)
    print "finished!"


def add_pagenum(canvas, doc):
    
    page_num = canvas.getPageNumber()
    text = "Page #%s" % page_num
    canvas.drawRightString(200*mm, 20*mm, text)




def createLineItems():
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
        ptext = "<font size=%s><b>%s</b></font>" % (font_size, text)
        p = Paragraph(ptext, centered)
        d.append(p)

    data = [d]

    line_num = 1

    formatted_line_data = []

    for each_rec in inp_data:
        line_data = [str(line_num),each_rec[0],each_rec[1],each_rec[2],each_rec[3]]

        for item in line_data:
            ptext = "<font size=%s>%s</font>" % (font_size-1, item)
            p = Paragraph(ptext, centered)
            formatted_line_data.append(p)
        data.append(formatted_line_data)
        formatted_line_data = []
        line_num += 1

    table = Table(data, colWidths=[75, 75, 150, 100, 100])

    self.story.append(table)









create_report()


























# #can = canvas.Canvas("report.pdf")
# def addPageNumber(canvas, doc):
#     """
#     Add the page number
#     """
#     page_num = canvas.getPageNumber()
#     text = "Page #%s" % page_num
#     canvas.drawRightString(200*mm, 20*mm, text)
#     
# def myFirstPage(can,doc):
#     #can.saveState()
#     can.setFont("Times-Roman", 10)
#     can.drawString(65,720,"Finance report for : %s %s" %(datetime.date.today().strftime("%B"), datetime.date.today().strftime("%Y")))
#     can.drawString(65,705,'Report generated on : %s' %str("{:%d.%m.%Y}".format(datetime.datetime.now())))
#     can.setLineWidth(.2)
#     can.line(50,695,540,695)
#     can.setFont("Times-Roman", 10)
#     can.drawString(50,680,"Sr. No")
#     can.drawString(90,680,"Ref. No")
#     can.drawString(185,680,"Name")
#     can.drawString(355,680,"Expected Amount")
#     can.drawString(455,680,"Paid Amount")
#     
#     line_start = 660
#     
#     for ix,each_rec in enumerate(data):
#         print each_rec
#         print line_start
#         can.setFont("Times-Roman", 10)
#         can.drawString(55,line_start,str(ix+1))
#         can.drawString(95,line_start,each_rec[0])
#         can.drawString(160,line_start,each_rec[1])
#         can.drawString(380,line_start,each_rec[2])
#         can.drawString(480,line_start,each_rec[3])
#         line_start= line_start-20
#         if line_start==60:
#             line_start=770
#             #can.showPage()
#     #can.restoreState()
#         
# styles = getSampleStyleSheet()
# 
# doc = SimpleDocTemplate("temp_report.pdf")
# Story = [Spacer(1,3.35*inch)]
# Story.append(PageBreak())
# style = styles["Normal"]
# 
# doc.build(Story, onFirstPage=myFirstPage, onLaterPages=addPageNumber)

