# from Tkinter import *
# 
# def select_all():
#     lb.select_set(0, END)
# 
# def select_none():
#     lb.selection_clear(0, END)
#     
# def print_selection():
#     items = map(int,lb.curselection())
#     if items:
#         for item in items:
#             print item
#     else:
#         print "Nothing to print"
# 
# root = Tk()
# lb = Listbox(root, selectmode=MULTIPLE)
# for i in range(10): 
#     lb.insert(END, i)
#     #ent.insert(0,END)
#     Entry(root)
# lb.grid(row=0)
# #ent.pack()
# Button(root, text='select all', command=select_all).grid(row=1)
# Button(root, text='select none', command=select_none).grid(row=2)
# Button(root, text='print selected', command=print_selection).grid(row=3)
# root.mainloop()
#!/usr/bin/env python

# ''' Create a grid of Tkinter Checkbuttons
# 
#     Each row permits a maximum of two selected buttons
# 
#     From http://stackoverflow.com/q/31410640/4014959
# 
#     Written by PM 2Ring 2015.07.15
# '''
# 
# import tkinter as tk
# 
# class Example(tk.Frame):
#     def __init__(self, parent):
#         tk.Frame.__init__(self, parent)
#         b = tk.Button(self, text="Done!", command=self.upload_cor)
#         b.pack()
#         table = tk.Frame(self)
#         table.pack(side="top", fill="both", expand=True)
# 
#         data = (
#             (45417, "rodringof", "CSP L2 Review", 0.000394, "2014-12-19 10:08:12", "2014-12-19 10:08:12"),
#             (45418, "rodringof", "CSP L2 Review", 0.000394, "2014-12-19 10:08:12", "2014-12-19 10:08:12"),
#             (45419, "rodringof", "CSP L2 Review", 0.000394, "2014-12-19 10:08:12", "2014-12-19 10:08:12"),
#             (45420, "rodringof", "CSP L2 Review", 0.000394, "2014-12-19 10:08:12", "2014-12-19 10:08:12"),
#             (45421, "rodringof", "CSP L2 Review", 0.000394, "2014-12-19 10:08:12", "2014-12-19 10:08:12"),
#             (45422, "rodringof", "CSP L2 Review", 0.000394, "2014-12-19 10:08:12", "2014-12-19 10:08:12"),
#             (45423, "rodringof", "CSP L2 Review", 0.000394, "2014-12-19 10:08:12", "2014-12-19 10:08:12"),
#         )
# 
#         self.widgets = {}
#         row = 0
#         for rowid, reviewer, task, num_seconds, start_time, end_time in (data):
#             row += 1
#             self.widgets[rowid] = {
#                 "rowid": tk.Label(table, text=rowid),
#                 "reviewer": tk.Label(table, text=reviewer),
#                 "task": tk.Label(table, text=task),
#                 "num_seconds_correction": tk.Entry(table),
#                 "num_seconds": tk.Label(table, text=num_seconds),
#                 "start_time": tk.Label(table, text=start_time),
#                 "end_time": tk.Label(table, text=start_time)
#             }
# 
#             self.widgets[rowid]["rowid"].grid(row=row, column=0, sticky="nsew")
#             self.widgets[rowid]["reviewer"].grid(row=row, column=1, sticky="nsew")
#             self.widgets[rowid]["task"].grid(row=row, column=2, sticky="nsew")
#             self.widgets[rowid]["num_seconds_correction"].grid(row=row, column=3, sticky="nsew")
#             self.widgets[rowid]["num_seconds"].grid(row=row, column=4, sticky="nsew")
#             self.widgets[rowid]["start_time"].grid(row=row, column=5, sticky="nsew")
#             self.widgets[rowid]["end_time"].grid(row=row, column=6, sticky="nsew")
# 
#         table.grid_columnconfigure(1, weight=1)
#         table.grid_columnconfigure(2, weight=1)
#         # invisible row after last row gets all extra space
#         table.grid_rowconfigure(row+1, weight=1)
# 
#     def upload_cor(self):
#         for rowid in sorted(self.widgets.keys()):
#             entry_widget = self.widgets[rowid]["num_seconds_correction"]
#             new_value = entry_widget.get()
#             print("%s: %s" % (rowid, new_value))
# 
# if __name__ == "__main__":
#     root = tk.Tk()
#     Example(root).pack(fill="both", expand=True)
#     root.mainloop()
# 
# from Tkinter import *
# 
# rows = ["A few lines", "of text", "for our example"]
# def callback(row):
#     print "you picked row # %s which has this data: %s" % (row, rows[row])
# 
# rows = ["A few lines", "of text", "for our example"]
# root = Tk()
# t = Text(root)
# t.pack()
# 
# t.insert(END, '\n'.join(rows))
# for i in range(len(rows)):
#     line_num = i + 1 # Tkinter text counts from 1, not zero
#     tag_name = "tag_%s" % line_num
#     t.tag_add(tag_name, "%s.0" % line_num, "%s.end" % line_num)
#     t.tag_bind(tag_name, "<Button-1>", lambda e, row=i: callback(row))
# 
# root.mainloop()

import Tkinter as tk

rows = ["A few lines", "of text", "for our example","A few lines", "of text", "for our example","A few lines", "of text", "for our example","A few lines", "of text", "for our example",
        "A few lines", "of text", "for our example","A few lines", "of text", "for our example","A few lines", "of text", "for our example","A few lines", "of text", "for our example"]
def callback(event):
    lb=event.widget
    # http://www.pythonware.com/library/tkinter/introduction/x5453-patterns.htm
    # http://www.pythonware.com/library/tkinter/introduction/x5513-methods.htm
    items = lb.curselection()
    try: items = map(int, items)
    except ValueError: pass
    idx=items[0]
    print(idx,rows[idx])       
root = tk.Tk()
scrollbar = tk.Scrollbar(root, orient="vertical")
lb = tk.Listbox(root, width=50, height=20,
                yscrollcommand=scrollbar.set)
scrollbar.config(command=lb.yview)
scrollbar.pack(side="right", fill="y")
lb.pack(side="left",fill="both", expand=True)
for row in rows:
    lb.insert("end", row)
    # http://www.pythonware.com/library/tkinter/introduction/events-and-bindings.htm
    lb.bind('<ButtonRelease-1>',callback)
root.mainloop()