import os, sys
import Tkinter as tk
import tkFileDialog, tkMessageBox
import win32ui, win32con, win32print

background_color='PeachPuff3'
global all_address
all_address = []

class AddressBook(tk.Frame):
    def __init__(self, master):
        self.master = master
        self.first_frame = tk.Frame(self.master, bg=background_color)
        self.first_frame.grid(row=5, column=0)
        self.file_name = "address_book.db"
        self.dir_path = ''
        self.first_page()

    def first_page(self):
        
        self.first_frame.destroy()
        self.first_frame = tk.Frame(self.master, bg=background_color)
        self.first_frame.grid(row=5, column=0)
        
        l1 = tk.Label(self.first_frame, text="Sender Details", bg=background_color)
        l1.grid(row=0, column=0, columnspan=2)
        l2 = tk.Label(self.first_frame, text="Receiver Details", bg=background_color)
        l2.grid(row=0, column=4, pady=10)
        
        s_l1 = tk.Label(self.first_frame, text="Name", bg=background_color)
        s_l1.grid(row=1, column=0)
        self.s_name = tk.Entry(self.first_frame, width=20)
        self.s_name.grid(row=1, column=1, padx=5)
        
        s_l2 = tk.Label(self.first_frame, text="Street, No", bg=background_color)
        s_l2.grid(row=2, column=0, padx=10)
        self.s_street = tk.Entry(self.first_frame, width=20)
        self.s_street.grid(row=2, column=1)
        
        s_l3 = tk.Label(self.first_frame, text="PostCode", bg=background_color)
        s_l3.grid(row=3, column=0)
        self.s_postcode = tk.Entry(self.first_frame, width=6)
        self.s_postcode.grid(row=3, column=1, sticky=tk.W, padx=5)
        
        s_l4 = tk.Label(self.first_frame, text="City", bg=background_color)
        s_l4.grid(row=4, column=0)
        self.s_city = tk.Entry(self.first_frame, width=20)
        self.s_city.grid(row=4, column=1, padx=5)
        
        r_l1 = tk.Label(self.first_frame, text="Name", bg=background_color)
        r_l1.grid(row=1, column=3)
        self.r_name = tk.Entry(self.first_frame, width=20)
        self.r_name.grid(row=1, column=4, padx=5)
        
        r_l2 = tk.Label(self.first_frame, text="Street, No", bg=background_color)
        r_l2.grid(row=2, column=3)
        self.r_street = tk.Entry(self.first_frame, width=20)
        self.r_street.grid(row=2, column=4, padx=5)
        
        r_l3 = tk.Label(self.first_frame, text="PostCode", bg=background_color)
        r_l3.grid(row=3, column=3)
        self.r_postcode = tk.Entry(self.first_frame, width=6)
        self.r_postcode.grid(row=3, column=4, padx=5, sticky=tk.W)
        
        r_l4 = tk.Label(self.first_frame, text="City", bg=background_color)
        r_l4.grid(row=4, column=3)
        self.r_city = tk.Entry(self.first_frame, width=20)
        self.r_city.grid(row=4, column=4, padx=5)

        b2 = tk.Button(self.first_frame,text="Print", command=self.validate_fields, height = 1, width = 15)
        b2.grid(row=9,column=1, padx=5, pady=25)
        b4 = tk.Button(self.first_frame,text="Quit", command=self.exit_prog, height = 1, width = 15)
        b4.grid(row=9,column=2, padx=5, pady=25)
        
        self.footer = tk.Label(self.first_frame, text="", fg="red", bg=background_color)
        self.footer.grid(row=10, column=1, columnspan=10)
    
    def exit_prog(self):
        if tkMessageBox.askokcancel("Quit", "Do you want to quit?"):
            root.destroy()
    
    def validate_fields(self):
        
        if self.s_name.index("end") == 0:
            self.s_name.config(bg="red")
            s_name_valid = False
        else:
            self.s_name.config(bg="white")
            s_name_valid = True
        if self.s_street.index("end") == 0:
            self.s_street.config(bg="red")
            s_street_valid = False
        else:
            self.s_street.config(bg="white")
            s_street_valid = True
        if self.s_postcode.index("end") == 0:
            self.s_postcode.config(bg="red")
            s_postcode_valid = False
        else:
            self.s_postcode.config(bg="white")
            s_postcode_valid = True
        if self.s_city.index("end") == 0:
            self.s_city.config(bg="red")
            s_city_valid = False
        else:
            self.s_city.config(bg="white")
            s_city_valid = True
        
        if self.r_name.index("end") == 0:
            self.r_name.config(bg="red")
            r_name_valid = False
        else:
            self.r_name.config(bg="white")
            r_name_valid = True
        if self.r_street.index("end") == 0:
            self.r_street.config(bg="red")
            r_street_valid = False
        else:
            self.r_street.config(bg="white")
            r_street_valid = True
        if self.r_postcode.index("end") == 0:
            self.r_postcode.config(bg="red")
            r_postcode_valid = False
        else:
            self.r_postcode.config(bg="white")
            r_postcode_valid = True
        if self.r_city.index("end") == 0:
            self.r_city.config(bg="red")
            r_city_valid = False
        else:
            self.r_city.config(bg="white")
            r_city_valid = True    
        
        if s_name_valid and s_street_valid and s_postcode_valid and s_city_valid and r_name_valid and r_street_valid and r_postcode_valid and r_city_valid:
            self.footer.config(text="")
            self.print_records()
        else:
            self.footer.config(text="Please check fields marked in red!", fg="red")
            
    def print_records(self):
        
        dc = win32ui.CreateDC()
        dc.CreatePrinterDC(win32print.GetDefaultPrinter())
        dc.SetMapMode(win32con.MM_TWIPS) # 1440 per inch
        scale_factor = 20 # i.e. 20 twips to the point
        dc.StartDoc("description")
        pen = win32ui.CreatePen(0, int(scale_factor), 0L)
        dc.SelectObject(pen)
        font = win32ui.CreateFont({
            "name": "Lucida Console",
            "height": int(scale_factor * 7),
            "weight": 400,
        })
        
        dc.SelectObject(font)
        dc.TextOut(scale_factor * 72,
            -1 * scale_factor * 120,
            self.s_name.get())
        
        dc.TextOut(scale_factor * 72,
            -1 * scale_factor * 131,
            self.s_street.get() + ', ' + self.s_postcode.get() + ' ' + self.s_city.get())
        
        dc.TextOut(scale_factor * 72,
            -1 * scale_factor * 140,
            "-------------------------------------------")
        
        font = win32ui.CreateFont({
            "name": "Lucida Console",
            "height": int(scale_factor * 10),
            "weight": 400,
        })
        dc.SelectObject(font)
        
        dc.TextOut(scale_factor * 72,
            -1 * scale_factor * 155,
            self.r_name.get())
        
        dc.TextOut(scale_factor * 72,
            -1 * scale_factor * 170,
            self.r_street.get() + ',')
        
        dc.TextOut(scale_factor * 72,
            -1 * scale_factor * 185,
            self.r_postcode.get() + ' ' + self.r_city.get())
        dc.EndDoc()
        
        print type(self.r_postcode.get())

root = tk.Tk()
root.geometry("600x350+400+400")
root.title("Address Formatter")
root.configure(background=background_color)
lf = AddressBook(root)
root.protocol("WM_DELETE_WINDOW", lf.exit_prog)
try:
    root.mainloop()
finally:
    import os
    print all_address
# # kill this process with taskkill
    current_pid = os.getpid()   
    os.system("taskkill /pid %s /f" % current_pid)

