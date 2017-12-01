import Tkinter as tk
import tkFileDialog, tkMessageBox
import sqlite3

background_color='PeachPuff3'

class ThreePinGUI():
    
    def __init__(self, master):
        self.master = master
        self.first_frame = tk.Frame(self.master, bg=background_color)
        self.first_frame.grid(row=5, column=0)
        self.check_database()
        self.lock_screen()
    
    def check_if_3_pin_enabled(self):
        query="Select * from pins"
        self.cur.execute(query)   
        recs = self.cur.fetchall()
        if len(recs) == 0:
            return False
        elif len(recs) == 3:
            self.data = recs
            return True    
    
    def check_invalid_tries(self):
        query="Select * from failed_tries"
        self.cur.execute(query)   
        recs = self.cur.fetchall()
        self.invalid_retries = recs
        if len(recs) == 2:
            return True
        else:
            return False 
        
    def check_database(self):
        self.conn = sqlite3.connect("pin_database.db")
        self.cur = self.conn.cursor()
        self.cur.execute("Create table IF NOT EXISTS pins (Id INT, pin INT, last_used TEXT )")
        self.conn.commit()
        self.cur.execute("Create table IF NOT EXISTS failed_tries (pin_id INT)")
        self.conn.commit()
    
    def lock_screen(self):
        self.first_frame.destroy()
        self.first_frame = tk.Frame(self.master, bg=background_color)
        self.first_frame.grid(row=5, column=0)
        frame = tk.Frame(self.first_frame, width=200, height=250, bg="light blue")
        frame.grid(row=0, column=0, padx=50,pady=30,columnspan=2)
        self.three_pin_enabled = self.check_if_3_pin_enabled() 
        if self.three_pin_enabled:
            invalid_retry_limit_reached = self.check_invalid_tries()
            if invalid_retry_limit_reached:
                tkMessageBox.showwarning("Unlock", "PIN unlock limit reached, Unlock using your ID")
                self.unlocked_id_screen()
            else:
                pin,pin_id,last_pin_id = self.pin_to_ask() 
                
                self.pin_label = tk.Label(self.first_frame, text="Enter PIN (%d) to unlock!" %int(pin_id+1))
                self.pin_label.grid(row=2,column=0, pady=10)
                self.user_pin = tk.Entry(self.first_frame, width=10)
                self.user_pin.grid(row=2,column=1)
                
                unlock_but = tk.Button(self.first_frame, text="Unlock", command= lambda: self.unlock_screen_with_pin(self.user_pin.get(), pin,pin_id,last_pin_id), width=12, height=1)
                unlock_but.grid(row=3,column=0,pady=2)
        else:
            unlock_but = tk.Button(self.first_frame, text="Unlock", command=self.unlocked_screen, width=12, height=1)
            unlock_but.grid(row=3,column=0,pady=2,columnspan=2)
    
    def unlocked_id_screen(self):
        self.first_frame.destroy()
        self.first_frame = tk.Frame(self.master, bg=background_color)
        self.first_frame.grid(row=5, column=0)
        frame = tk.Frame(self.first_frame, width=200, height=250)
        frame.grid(row=0, column=0, padx=50,pady=30,columnspan=5)
        unlock_but = tk.Button(self.first_frame, text="Submit", command=self.unlock_using_id, width=12, height=1)
        unlock_but.grid(row=3,column=0,pady=2,columnspan=2)
        
    def unlock_using_id(self):    
        try:
            query="Delete from failed_tries"
            self.cur.execute(query)   
            self.conn.commit()
        except Exception, msg:
            tkMessageBox.showerror("ID unlock", msg)
        else:
            self.lock_screen()
    
    def pin_to_ask(self):
        next_pin=0
        first_use_check=0
        last_used_pin=0
        for each_rec in self.data:
            last_used = each_rec[2].encode("utf-8")
            print each_rec, last_used
            if last_used == "Yes":
                last_used_pin = each_rec[0]
                next_pin = last_used_pin + 1
                if next_pin >= 3 :
                    next_pin = 0
                break
            else:
                first_use_check += 1
                
        if first_use_check == 3:
            return self.data[next_pin][1], next_pin, ""
        else:
            return self.data[next_pin][1], next_pin, last_used_pin
    
    def unlock_screen_with_pin(self, user_pin, pin, pin_id, last_pin_id):
        try:
            user_pin = int(user_pin)
        except Exception, msg:
            self.update_invalid_retries(pin, pin_id, last_pin_id)
            tkMessageBox.showerror("Unlock", "Invalid PIN, please retry!")
            self.lock_screen()
        else:
            if user_pin == pin:
                print "Success"
                if not last_pin_id:   
                    args = [" ",0]
                else:
                    args = [" ",last_pin_id]
                try:
                    sql = "Update pins SET last_used=? where Id=?"
                    self.cur.execute(sql,args)
                    self.conn.commit()
        #              
                    sql = "Update pins SET last_used=? where Id=?"
                    args = ["Yes",pin_id]
                    self.cur.execute(sql,args)
                    self.conn.commit()
                except Exception, msg:
                    tkMessageBox.showerror("Unlock", msg)
                else:
                    self.unlocked_screen()
            else:
                self.update_invalid_retries(pin, pin_id, last_pin_id)
                tkMessageBox.showerror("Unlock", "Invalid PIN, please retry!")
                self.lock_screen()
    
    def validate_user_entered_pin(self, user_pin):
        if not user_pin.isdigit():
            tkMessageBox.showerror("PIN-1", "Enter only numbers")
        elif len(user_pin) < 4 or len(user_pin) > 4:
            tkMessageBox.showerror("PIN-1", "PIN length should be only 4 digits")
        else:
            return int(user_pin)
            
    def update_invalid_retries(self, pin, pin_id,last_pin_id):
        sql = "insert into failed_tries VALUES (?) "
        args=[pin_id]
        self.cur.execute(sql,args)
        self.conn.commit()
        
        args = [" ",last_pin_id]
        sql = "Update pins SET last_used=? where Id=?"
        self.cur.execute(sql,args)
        self.conn.commit()
                
        sql = "Update pins SET last_used=? where Id=?"
        args = ["Yes",pin_id]
        self.cur.execute(sql,args)
        self.conn.commit()

    def unlocked_screen(self):
        self.first_frame.destroy()
        self.first_frame = tk.Frame(self.master, bg=background_color)
        self.first_frame.grid(row=5, column=0)
        frame = tk.Frame(self.first_frame, bg=background_color)
        frame.grid(row=1, column=1, padx=50,pady=50,columnspan=2)
        set_3_pin_but = tk.Button(frame, text="Enable 3 PIN", command=self.enable_3_pin, width=12,height=1)
        set_3_pin_but.grid(row=3,column=0,pady=2,padx=10)
        unset_3_pin_but = tk.Button(frame, text="Disable 3 PIN", command=self.disable_3_pin, width=12,height=1)
        unset_3_pin_but.grid(row=3,column=1,pady=2)
        lock_scr = tk.Button(frame, text="Lock Screen", command=self.lock_screen, width=12,height=1)
        lock_scr.grid(row=4,column=0,pady=10)
        
    def enable_3_pin(self):
        three_pin_enabled = self.check_if_3_pin_enabled()
        if three_pin_enabled:
            tkMessageBox.showinfo("Three PIN", "Three PIN authentication is already enabled!")
        else:
            self.first_frame.destroy()
            self.first_frame = tk.Frame(self.master, bg=background_color)
            self.first_frame.grid(row=5, column=0)
            frame = tk.Frame(self.first_frame, bg=background_color)
            frame.grid(row=1, column=1, padx=50,pady=50,columnspan=2)
            main_label = tk.Label(frame, text="Enter 3 different PINS, each of length 4!", relief=tk.RAISED, borderwidth=2, bg="grey")
            main_label.grid(row=0,column=0, padx=5, pady=10, columnspan=5)
            pin1_label = tk.Label(frame, text="PIN-1", relief=tk.RAISED, borderwidth=2, bg="grey")
            pin1_label.grid(row=1,column=0, padx=5, pady=2)
            pin_1 = tk.Entry(frame, width=12, relief=tk.SUNKEN, borderwidth=2)
            pin_1.grid(row=1,column=1)
            pin2_label = tk.Label(frame, text="PIN-2",relief=tk.RAISED, borderwidth=2, bg="grey")
            pin2_label.grid(row=2,column=0, padx=5, pady=2)
            pin_2 = tk.Entry(frame, width=12, relief=tk.SUNKEN, borderwidth=2)
            pin_2.grid(row=2,column=1)
            pin3_label = tk.Label(frame, text="PIN-3",relief=tk.RAISED, borderwidth=2, bg="grey")
            pin3_label.grid(row=3,column=0, padx=5, pady=2)
            pin_3 = tk.Entry(frame, width=12, relief=tk.SUNKEN, borderwidth=2)
            pin_3.grid(row=3,column=1)
            lock_scr = tk.Button(frame, text="Submit", command= lambda: self.setup_3_pin(pin_1.get(),pin_2.get(),pin_3.get()),
                                  width=12,height=1)
            lock_scr.grid(row=4,column=0,pady=25, columnspan=5)
        
    def setup_3_pin(self,pin1,pin2,pin3):
        if pin1 == "":
            tkMessageBox.showerror("Invalid PIN", "PIN-1 cannot be empty!")
        elif pin2 == "":
            tkMessageBox.showerror("Invalid PIN", "PIN-2 cannot be empty!")
        elif pin3 == "":
            tkMessageBox.showerror("Invalid PIN", "PIN-3 cannot be empty!")
        elif not pin1.isdigit():
            tkMessageBox.showerror("Invalid PIN", "Enter only numbers PIN-1")
        elif not pin2.isdigit():
            tkMessageBox.showerror("Invalid PIN", "Enter only numbers PIN-2")
        elif not pin3.isdigit():
            tkMessageBox.showerror("Invalid PIN", "Enter only numbers for PIN-3")
        elif len(pin1) < 4 or len(pin1) > 4:
            tkMessageBox.showerror("Invalid PIN", "Length of PIN-1 should be only 4 digits")
        elif len(pin2) < 4 or len(pin2) > 4:
            tkMessageBox.showerror("Invalid PIN", "Length of PIN-2 should be only 4 digits")
        elif len(pin3) < 4 or len(pin3) > 4:
            tkMessageBox.showerror("Invalid PIN", "Length of PIN-3 should be only 4 digits")
        elif pin1 == pin2 or pin1 == pin3 or pin2 == pin3:
            tkMessageBox.showerror("Invalid PIN", "PINs cannot be same! Enter different PINs")
        else:
            sql1 = "insert into pins VALUES (?,?,?) "
            args1 = [0,pin1,""]
            
            sql2 = "insert into pins VALUES (?,?,?) "
            args2 = [1,pin2,""]    
            
            sql3 = "insert into pins VALUES (?,?,?) "
            args3 = [2,pin3,""]
            try:
                self.cur.execute(sql1,args1)
                self.cur.execute(sql2,args2)
                self.cur.execute(sql3,args3)
                self.conn.commit()
            except Exception, msg:
                tkMessageBox.showerror("Error", "%s occurred while adding new record! " %str(msg))
            else:
                self.lock_screen()
        
    def disable_3_pin(self):
        three_pin_enabled = self.check_if_3_pin_enabled()
        if three_pin_enabled:
            try:
                query="Delete from pins"
                self.cur.execute(query)   
                self.conn.commit()
            except Exception, msg:
                tkMessageBox.showerror("Disable 3 PIN", msg)
            else:
                tkMessageBox.showinfo("Three PIN", "Three PIN authentication disabled successfully!")
        else:
            tkMessageBox.showinfo("Three PIN", "Three PIN authentication is not enabled, cannot disable it")

        
import tkFont
root = tk.Tk()
my_font = tkFont.Font(family="Monaco", size=9)
root.geometry("300x400+400+400")
root.title("Three PIN")
root.configure(background=background_color)
try:
    ThreePinGUI(root)
    root.mainloop()
except Exception, msg:
    tkMessageBox.showerror("Error","%s occurred " %str(msg))