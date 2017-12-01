import sqlite3

class ThreePinAuth():
    
    def __init__(self):
        pass
    
    def check_database(self):
        self.conn = sqlite3.connect("pin_database.db")
        self.cur = self.conn.cursor()
        self.cur.execute("Create table IF NOT EXISTS pins (Id INT, pin INT, last_used TEXT )")
        self.conn.commit()
        
        self.cur.execute("Create table IF NOT EXISTS failed_tries (pin_id INT)")
        self.conn.commit()
        
        self.cur.execute("Create table IF NOT EXISTS account_details (email TEXT, password TEXT)")
        self.conn.commit()
    
    def check_if_3_pin_enabled(self):
        query="Select * from pins"
        self.cur.execute(query)   
        recs = self.cur.fetchall()
        if len(recs) == 0:
            print "3 pin authentication not enabled"
            print recs
        elif len(recs) == 3:
            print "3 pin authentication enabled"    
            self.data = recs
    
    def pin_to_ask(self):
        next_pin=0
        first_use_check=0
        last_used_pin=0
        
        self.check_if_3_pin_enabled()
        
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
    
    def ask_pin(self):
        
        pin,pin_id,last_pin_id = self.pin_to_ask()
        
        i=0
        while (i < 3):
            i+=1
            user_pin = int(raw_input("Enter your PIN (%d)" %int(pin_id+1)))
            if user_pin == pin:
                print "Success"
                if not last_pin_id:   
                    args = [" ",0]
                else:
                    args = [" ",last_pin_id]
                sql = "Update pins SET last_used=? where Id=?"
                self.cur.execute(sql,args)
                self.conn.commit()
    #              
                sql = "Update pins SET last_used=? where Id=?"
                args = ["Yes",pin_id]
                self.cur.execute(sql,args)
                self.conn.commit()
                break                 
            else:
                print "False PIN, Try again!"
        else:
            self.incorrect_tries(pin_id,last_pin_id)
            
    
    def incorrect_tries(self, pin_id,last_pin_id):
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
        
        self.check_invalid_tries()
        
        
    def check_invalid_tries(self):
        query="Select * from failed_tries"
        self.cur.execute(query)   
        recs = self.cur.fetchall()
        print recs
        if len(recs) == 2:
            print "2 invalid tries found, please use id to unlock!"
            self.use_id_to_unlock()
            print recs
        else:
            self.ask_pin()
        
    def use_id_to_unlock(self):
        print "Use ID"
        
        email = raw_input("Enter email id")
        password = raw_input("Enter Password")
        
        query = "Select * from account_details"
        self.cur.execute(query)
        recs = self.cur.fetchall() 
        
        if email == recs[0] and password == recs[1]:
        
            query="Delete from failed_tries"
            self.cur.execute(query)   
            self.conn.commit()
        else:
            print "Kuch toh gadbad hai!"
        
    
                    
    def main(self):
        self.check_database()
        self.check_invalid_tries()
        #self.ask_pin()


pin3 = ThreePinAuth()  
pin3.main()               