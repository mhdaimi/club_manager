# -*- coding: utf-8 -*-

from reportlab.pdfgen import canvas 
from reportlab.pdfbase import pdfmetrics 
from reportlab.pdfbase.ttfonts import TTFont 

pdfmetrics.registerFont(TTFont('Hebrew', 'ArialHB.ttf'))

def hello(c):
    c.setFont("Hebrew", 14)
    c.drawString(10,10, u"מה שלומך".encode('utf-8'))

c = canvas.Canvas("hello.pdf") 
hello(c) 
c.showPage()
c.save()