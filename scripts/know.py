from tkinter import *
from tkinter.ttk import Progressbar
from time import sleep

def RunThat():
   global pb
   for i in range(0,101,10):
      pb["value"]=i
      pb.update()  #this works
      sleep(0.1)
      print(i)

master=Tk()

Label(master,text="Doing something...").grid(row=0,column=0,pady=10)
Grid.columnconfigure(master,0,weight=1)
Label(master,text="Progress:",justify=LEFT).grid(row=1,column=0,padx=25,pady=2,sticky=W+S)
Grid.rowconfigure(master,1,weight=1)
pb=Progressbar(master)
pb.grid(row=2,column=0,padx=25,pady=2,sticky=W+E)   
pb["maximum"]=100
pb["value"]=0  
Button(master,text="Do it",command=RunThat).grid(row=3,column=0,padx=5,pady=5)

master.mainloop()