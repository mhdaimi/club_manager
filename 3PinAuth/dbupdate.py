import sqlite3

conn = sqlite3.connect("pin_database.db")
cur = conn.cursor()

sql = "select * from pins"


#sql = "insert into account_details VALUES (?,?) "
#args=["mhdaimi@hotmail.com","Smfs@1208"]

cur.execute(sql)
print cur.fetchall()