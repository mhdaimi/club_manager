import os, sys
import Tkinter as tk
import tkFileDialog, tkMessageBox

background_color='PeachPuff3'
global all_address
all_address = []

class ScrollableFrame(tk.Frame):
    def __init__(self, master, **kwargs):
        tk.Frame.__init__(self, master, **kwargs)

        # create a canvas object and a vertical scrollbar for scrolling it
        self.vscrollbar = tk.Scrollbar(self, orient=tk.VERTICAL)
        self.vscrollbar.pack(side='right', fill="y",  expand="false")
        self.canvas = tk.Canvas(self,
                                bg=background_color, bd=0,
                                height=300,
                                width=680,
                                highlightthickness=0,
                                yscrollcommand=self.vscrollbar.set)
        self.canvas.pack(side="left", fill="both", expand="true")
        self.vscrollbar.config(command=self.canvas.yview)

        # reset the view
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)

        # create a frame inside the canvas which will be scrolled with it
        self.interior = tk.Frame(self.canvas, **kwargs)
        self.canvas.create_window(0, 0, window=self.interior, anchor="nw")

        self.bind('<Configure>', self.set_scrollregion)


    def set_scrollregion(self, event=None):
        """ Set the scroll region on the canvas"""
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

class AddressBook(tk.Frame):
    
    def __init__(self, master):
        self.master = master
#         self.first_frame = tk.Frame(self.master, bg=background_color)
#         self.first_frame.grid(row=5, column=0)
        self.file_name = "address_book.txt"
        self.dir_path = ''
        
        self.check_config_file()
        
    def check_config_file(self):
        
        if os.path.isfile(os.path.join(os.path.expanduser("~"), 'address_book.config')):
            with open (os.path.join(os.path.expanduser("~"), 'address_book.config'), 'r') as fh:
                dir_path = fh.readlines()
                #print dir_path[0]
                self.dir_path = dir_path[0]
                self.first_page(from_init=True)
        else:
            
            dir_name = tkFileDialog.askdirectory(parent=self.master,title='Please select a directory')
            if dir_name:
                with open (os.path.join(os.path.expanduser("~"), 'address_book.config'), 'w') as fh:
                    fh.writelines(dir_name)
                    self.dir_path = dir_name
                    self.first_page(from_init=True)
            else:
                tkMessageBox.showerror("Directory Not Set", "Restart the program and select a directory to continue!")
                root.destroy()
        
    def first_page(self, check='', from_init=''):
        global all_address
        all_address=[]
        if check:
            self.checkbox_pane.destroy()
        
        if not from_init:
            self.first_frame.destroy()
        self.first_frame = tk.Frame(self.master, bg=background_color)
        self.first_frame.grid(row=5, column=0)
        
        b3 = tk.Button(self.first_frame,text="Add", command=self.add_address, height = 1, width = 15)
        b3.grid(row=8,column=1,padx=5, pady=50)
        b1 = tk.Button(self.first_frame,text="Manage", command=self.manage_address, height = 1, width = 15)
        b1.grid(row=8,column=2, padx=5, pady=25)
        b2 = tk.Button(self.first_frame,text="Print", command=self.print_records, height = 1, width = 15)
        b2.grid(row=9,column=1, padx=5, pady=25)
        b4 = tk.Button(self.first_frame,text="Quit", command=self.exit_prog, height = 1, width = 15)
        b4.grid(row=9,column=2, padx=5, pady=25)
        
        l1 = tk.Label(self.first_frame,text="Your database file is stored at: %s" %(self.dir_path + '/' + self.file_name))
        l1.grid(row=25, columnspan=5, padx=5, pady=300)
    
    def exit_prog(self):
        
        root.destroy()
        
            
    def add_address(self):
        self.first_frame.destroy()
        self.first_frame = tk.Frame(self.master, bg=background_color)
        self.first_frame.grid(row=5, column=0)
        
        label_1 = tk.Label(self.first_frame, text="Name: ", bg=background_color)
        #label_2 = Label(self.first_frame, text="Last Name: ", bg=background_color)
        label_3 = tk.Label(self.first_frame, text="Street, House No : ", bg=background_color)
        #label_4 = Label(self.first_frame, text="House No : ", bg=background_color)
        label_5 = tk.Label(self.first_frame, text="PostCode, City : ", bg=background_color)
        #label_6 = Label(self.first_frame, text="Post Code : ", bg=background_color)
        label_7 = tk.Label(self.first_frame, text="Phone No : ", bg=background_color)
        
        self.f_name = tk.Entry(self.first_frame, width = 40)
        #self.l_name = Entry(self.first_frame, width = 40)
        self.street = tk.Entry(self.first_frame, width = 40)
        #self.h_no = Entry(self.first_frame, width = 40)
        self.city = tk.Entry(self.first_frame, width = 40)
        #self.post_code = Entry(self.first_frame, width = 40)
        self.phone = tk.Entry(self.first_frame, width = 40)
        
        label_1.grid(row=0, sticky=tk.E, padx=25, pady=10)
        #label_2.grid(row=1, sticky=E, padx=25, pady=10)
        label_3.grid(row=2, sticky=tk.E, padx=25, pady=10)
        #label_4.grid(row=3, sticky=E, padx=25, pady=10)
        label_5.grid(row=4, sticky=tk.E, padx=25, pady=10)
        #label_6.grid(row=5, sticky=E, padx=25, pady=10)
        label_7.grid(row=6, sticky=tk.E, padx=25, pady=10)
        
        self.f_name.grid(row=0, column=1, padx=5, columnspan=2)
        #self.l_name.grid(row=1, column=1, columnspan=2)
        self.street.grid(row=2, column=1, columnspan=2)
        #self.h_no.grid(row=3, column=1, columnspan=2)
        self.city.grid(row=4, column=1, columnspan=2)
        #self.post_code.grid(row=5, column=1, columnspan=2)
        self.phone.grid(row=6, column=1, columnspan=2)
        
        button_1 = tk.Button(self.first_frame, text="Submit", command=self.add_record, height = 1, width = 15)
        button_1.grid(row=15, column=0, padx=5, pady=25)
        
        button_2 = tk.Button(self.first_frame, text="Reset", command=self.reset_fields, height = 1, width = 15)
        button_2.grid(row=15, column=1, padx=5, pady=25)
        
        button_3 = tk.Button(self.first_frame, text="Back", command=self.first_page, height = 1, width = 15)
        button_3.grid(row=15, column=2, padx=5, pady=25)
        
        self.footer = tk.Label(self.first_frame, text="", fg="red", bg=background_color)
        self.footer.grid(row=18, column=1, columnspan=10)
        
    def add_record(self):
        address_rec = "%s;%40s;%25s\n"
        
        if self.f_name.index("end") == 0:
            self.footer.config(text="First Name cannot be blank")
#         elif self.l_name.index("end") == 0:
#             self.footer.config(text="Last Name cannot be blank")
        elif self.street.index("end") == 0:
            self.footer.config(text="Street cannot be blank")
#         elif self.h_no.index("end") == 0:
#             self.footer.config(text="House number cannot be blank")
        elif self.city.index("end") == 0:
            self.footer.config(text="City cannot be blank")
#         elif self.post_code.index("end") == 0:
#             self.footer.config(text="Post Code cannot be blank")
        elif self.phone.index("end") == 0:
            self.footer.config(text="Telefone number cannot be blank")
        else:
            address = self.street.get() + "," +  self.city.get()
            
            rec = address_rec %(self.f_name.get(), address, self.phone.get())
            with open(self.dir_path + '/' + self.file_name, "a") as file_obj:
                file_obj.writelines(rec)
#             else:
#                 rec = address_rec %(self.f_name.get(), self.l_name.get(), address, self.phone.get())
#                 with open("C:/opt/address_book.txt", "w") as file_obj:
#                     file_obj.writelines(rec)
                    
            self.f_name.delete(0, tk.END)
            #self.l_name.delete(0, END)
            self.street.delete(0, tk.END)
            #self.h_no.delete(0, END)
            self.city.delete(0, tk.END)
            #self.post_code.delete(0, END)
            self.phone.delete(0, tk.END)
            
            self.footer.config(text="Record added successfully!", fg="green")
        
    def reset_fields(self):
        self.first_frame.destroy()
        self.add_address()
            
    def manage_address(self):
        self.first_frame.destroy()
        
        self.first_frame = tk.Frame(self.master, bg=background_color)
        self.first_frame.grid(row=5, column=0)
        
        self.display_recs('')
        
        self.f_name = tk.Entry(self.first_frame, width=40)
        self.l_name = tk.Entry(self.first_frame, width=40)
        self.phone = tk.Entry(self.first_frame, width=40)
        search_fields = ["Name", "City", "Phone"]
        self.var = tk.StringVar()
        rows=5
        self.radio_butt = []
        for i, field in zip(range(len(search_fields)),search_fields):
            rows += 1
            
            self.radio_butt.append(tk.Radiobutton(self.first_frame, text=field, variable=self.var, value=field, height = 1, width = 15, 
                                               command=self.selected, bg=background_color ))
            self.radio_butt[i].grid(row=rows, pady=5, column=0, padx=5)
        self.radio_butt[0].select()

        self.f_name.grid(row=6, column=1, columnspan=2)
        self.l_name.grid(row=7, column=1, columnspan=2)
        self.phone.grid(row=8, column=1, columnspan=2)
        self.phone.config(state="disabled")
        self.l_name.config(state="disabled")
        
        button_1 = tk.Button(self.first_frame, text="Search", command=self.fetch_data, height = 1, width = 15)
        button_1.grid(row=15, column=0, padx=5, pady=10)
        button_2 = tk.Button(self.first_frame, text="Start Page", command=lambda: self.first_page('from_manage'), height = 1, width = 15)
        button_2.grid(row=15, column=1, padx=5, pady=10)
        
        button_3 = tk.Button(self.first_frame, text="Delete", command=self.delete_recs, height = 1, width = 15)
        button_3.grid(row=15, column=2, padx=5, pady=10)
        
        self.search_page_footer = tk.Label(self.first_frame, text="", fg="red", bg=background_color)
        self.search_page_footer.grid(row=18, column=1, columnspan=10)
    
    
    def delete_recs(self):
        
        global all_address
        
        if not all_address:
            tkMessageBox.showinfo('Delete Records', "Please select record(s) to delete!")
        else:
            var = tkMessageBox.askyesno('Delete Records', "Are you sure you want to delete %s record(s)?" %(str(len(all_address))))
    
            if var:
                with open(self.dir_path + '/' + self.file_name, "r") as read_obj:
                    all_lines = read_obj.readlines()
                
                for rec_to_del in all_address:
                    #print all_lines.count(rec_to_del+'\n'), rec_to_del
                    all_lines.remove(rec_to_del+'\n')
                
                with open(self.dir_path + '/' + self.file_name, "w") as write_obj:
                    write_obj.writelines(all_lines)
                 
            else:
                pass
        
        
    def fetch_data(self):
        result_list = []
        
#         if self.var.get() == "Name" and self.f_name.index("end") == 0:
#             self.search_page_footer.config(text="Name cannot be blank")
#         elif self.var.get() == "City" and self.l_name.index("end") == 0:
#             self.search_page_footer.config(text="City cannot be blank")
#         elif self.var.get() == "Phone" and self.phone.index("end") == 0:
#             self.search_page_footer.config(text="Phone number cannot be blank")
        with open(self.dir_path + '/' + self.file_name, "r") as read_obj:
            all_lines = read_obj.readlines()

        if (self.var.get() == "Name" and self.f_name.index("end") == 0) and \
            (self.var.get() == "City" and self.l_name.index("end") == 0) and \
            (self.var.get() == "Phone" and self.phone.index("end") == 0):
            
            result_list = all_lines
        else:
            self.checkbox_pane.destroy()
            self.search_page_footer.config(text="")
                    
            for each_line in all_lines:
                
                each_val = each_line.split(";")
                name = each_val[0]
                address = each_val[1]
                phone_num = each_val[2]
                if self.var.get() == "Name" and self.f_name.get().lower() in name.lower() :
                    #result_list += ("{:^15} {:^15} {:^30} {:>15}\n".format(*each_line.split(";")))
                    result_list.append(each_line)
                    continue
                elif self.var.get() == "City" and self.l_name.get().lower() in address.lower() :
                    #result_list += ("{:^15} {:^15} {:^30} {:>15}\n".format(*each_line.split(";")))
                    result_list.append(each_line)
                    continue
                elif self.var.get() == "Phone" and self.phone.get() in phone_num :
                    #result_list += ("{:^15} {:^15} {:^30} {:>15}\n".format(*each_line.split(";")))
                    result_list.append(each_line)
                    continue
        
            #self.show_results(result_list)
        self.display_recs(result_list)
    
    def selected(self):
        
        if self.var.get() == "City":
            self.l_name.config(state="normal")
            self.f_name.delete(0, 'end')
            self.phone.delete(0, 'end')
            self.f_name.config(state="disabled")
            self.phone.config(state="disabled")
        elif self.var.get() == "Phone":
            self.phone.config(state="normal")
            self.f_name.delete(0, 'end')
            self.l_name.delete(0, 'end')
            self.f_name.config(state="disabled")
            self.l_name.config(state="disabled")
        elif self.var.get() == "Name":
            self.f_name.config(state="normal")
            self.phone.delete(0, 'end')
            self.l_name.delete(0, 'end')
            self.phone.config(state="disabled")
            self.l_name.config(state="disabled")
            
    
    def show_results(self, data):

        self.first_frame.destroy()
        self.first_frame = tk.Frame(self.master, bg=background_color)
        self.first_frame.grid(row=5, columnspan=10)
        self.create_labels()
        row=7
        column=0
        for i in range(len(data)):
            each_rec = data[i]
            vals = each_rec.split(";")
            for each_val in vals:
                label1 = tk.Label(self.first_frame, text=each_val , bg="brown",height = 2, width = 20, wraplength=100)
                label1.grid(row=row, column=column, pady=10, padx=5)
                column+=2
            row+=1
            column=0
        
        back_but = tk.Button(self.first_frame, text="Back", command=self.search_address, height = 2, width = 15)
        back_but.grid(row=50, sticky=tk.E)
            
    def print_records(self):
        global all_address
        all_address=[]
        self.first_frame.destroy()
        self.first_frame = tk.Frame(self.master, bg=background_color)
        self.first_frame.grid(row=50, column=10)
        
        with open(self.dir_path + '/' + self.file_name, "r") as read_obj:
            all_lines = read_obj.readlines()
        
        self.display_recs(all_lines)
            

        b1 = tk.Button(self.first_frame, text = 'Select All', command = self.select_all, height = 1, width = 15)
        b1.grid(row=50, column=11, padx=10, pady=10)
        b2 = tk.Button(self.first_frame, text = 'De-select', command = self.deselect_all, height = 1, width = 15)
        b2.grid(row=50, column=12, padx=10)
        b3 = tk.Button(self.first_frame, text = 'Print', command = self.prepare_to_printer, height = 1, width = 15)
        b3.grid(row=50, column=13, padx=10)
        button_1 = tk.Button(self.first_frame, text="Start Page", command=lambda : self.first_page('from_print'), height = 1, width = 15)
        button_1.grid(row=51, column=11, padx=5, pady=5)
        
    
    def display_recs(self,data):
        
        self.checkbox_pane = ScrollableFrame(self.master, bg=background_color)
        self.checkbox_pane.grid(row=1, columnspan=25, padx=5, pady=10)
        
        self.cb = []
        self.cb_v = []  
        for ix, text in enumerate(data):
            self.cb_v.append(tk.IntVar())
            self.cb.append(tk.Checkbutton(self.checkbox_pane.interior, text=text.strip('\n'), variable=self.cb_v[ix], command=self.cb_checked, bg=background_color))
    
            self.cb[ix].grid(row=ix, column=0, sticky='w', padx=5, pady=5, columnspan=10)
        
        
    def prepare_to_printer(self):
        global all_address
        
        if not all_address:
            tkMessageBox.showinfo('Print Records', "Please select record(s) to print!")
        else:
            var = tkMessageBox.askyesno('Print Records', "Are you sure you want to print %s record(s)?" %(str(len(all_address))))
        
            if var:
                for each_record in all_address:
                    
                    each_val = each_record.split(';')
                    address = each_val[1].split(',')
                    raw_data = "\n\n\n\t\t%s\n\t\t%s,\n\t\t%s" %(each_val[0], address[0].strip(), address[1])
                    #print raw_data
                    
                    self.send_to_printer(raw_data)
            else:
                pass
            
    def send_to_printer(self,raw_data):
        import win32ui
        import win32print
        import win32con
        
        INCH = 2000
        try:
            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC (win32print.GetDefaultPrinter ())
            hDC.StartDoc("Test doc")
            hDC.StartPage()
            hDC.SetMapMode (win32con.MM_TWIPS)
            hDC.DrawText (raw_data, (-2, INCH * -1, INCH * 12, INCH * -4), win32con.DT_EXPANDTABS)
        finally:
            hDC.EndPage()
            hDC.EndDoc()
            hDC.DeleteDC()
        
    def cb_checked(self):
        global all_address
        all_address=[]

        for ix, item in enumerate(self.cb):
            if self.cb_v[ix].get():
                all_address.append(item['text'])
                
    
    def select_all(self):
        global all_address
        all_address = []
        for i in self.cb:
            i.select()
            all_address.append(i.cget("text"))
    
    def deselect_all(self):
        global all_address
        all_address=[]
        for i in self.cb:
            i.deselect()
        
    def create_text(self):
        self.text_wid = tk.Text(self.first_frame)
        self.text_wid.grid(row=7, column=0,  padx=5, pady=10)
        
    def create_labels(self):
         
            self.label1 = tk.Label(self.first_frame, text='First Name : ' , bg="brown",height = 1, width = 15)
            self.label1.grid(row=6, column=0, pady=10, padx=5)
            self.label2 = tk.Label(self.first_frame, text='Last Name : ' , bg="yellow",height = 1, width = 15)
            self.label2.grid(row=6, column=2, pady=10, padx=1)
            self.label3 = tk.Label(self.first_frame, text= 'Address :' , bg="green",height = 1, width = 15)
            self.label3.grid(row=6, column=4, pady=10, padx=1)
            self.label4 = tk.Label(self.first_frame, text= 'Contact No. : ' , bg="orange",height = 1, width = 15)
            self.label4.grid(row=6, column=6, pady=10, padx=1)


root = tk.Tk()
root.geometry("750x500+400+400")
root.title("Address Book")
root.configure(background=background_color)
lf = AddressBook(root)
root.mainloop()
#print all_address