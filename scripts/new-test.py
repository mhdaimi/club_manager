import Tkinter as tk
import sqlite3, os

master = tk.Tk()
master.title("Select Groups")

class MultiCB():
    
    def __init__(self, master):

        self.master = master
        self.max_num = 3
        self.check_but_list=[]
        self.label_1_list = []
        self.label_2_list = []
        self.dict_for = {}
        self.cb_v = []
        self.db_file = "test_db.db"
        self.conn=sqlite3.connect(self.db_file)
        self.cur = self.conn.cursor()
        
        self.cur.execute("""
                        CREATE table IF NOT EXISTS address(name text, address text, phone real, email text)
            """)
        self.conn.commit()
        
        self.main_prog()
    

    def add_record(self):
    
        self.cur.execute("""
                             INSERT into address VALUES (?,?,?,?)
                         """, [self.f_name.get(), self.street.get(), int(self.house_no.get()), self.post_code.get() ])
        self.conn.commit()
        
    def show_record(self):
        
        self.cur.execute(""" SELECT * from address""")
        result = self.cur.fetchall()
        
        print result
        
        for each_result in result:
            print each_result
            print each_result[0]
            print each_result[1]
            print each_result[2]
            
            
    
    def main_prog(self):
        
        self.first_frame = tk.Frame(self.master)
        self.first_frame.grid(row=0)
        
        self.f_name = tk.Entry(self.first_frame, width = 40)
        self.f_name.grid(row=0, column=2, sticky=tk.W, columnspan=5)
        self.street = tk.Entry(self.first_frame, width = 15)
        self.street.grid(row=1, column=2, sticky=tk.W)
        self.house_no = tk.Entry(self.first_frame, width = 3)
        self.house_no.grid(row=1, column=4, sticky=tk.W)
        self.post_code = tk.Entry(self.first_frame, width = 6)
        self.post_code.grid(row=2, column=2, sticky=tk.W)
        
        button_1 = tk.Button(self.first_frame, text="Submit", command=self.add_record, height = 1, width = 15)
        button_1.grid(row=15, column=2, padx=5, pady=25)
        
        button_1 = tk.Button(self.first_frame, text="Show", command=self.show_record, height = 1, width = 15)
        button_1.grid(row=15, column=3, padx=5, pady=25)
                
        


MultiCB(master)
try:
    master.mainloop()
finally:
    pass