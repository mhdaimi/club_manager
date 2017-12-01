#!/usr/bin/python
# -*- coding: utf-8 -*-
import sqlite3, os, string

conn=sqlite3.connect("C:/opt/Address Book/verein_manager.db")
cur = conn.cursor()


    
sql = """INSERT into member_data VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?) """


for i in range(1,501,1):
    if i % 2 ==0:
        gender = "Male"
        name = "mohammad daimi"
        city = "frankfurt"
        pay=20
    else:
        gender = "Female"
        name = "Sheher Bano"
        city = u"Königstein im Taunus"
        pay=25.5
        
    parameters = [i, str(i).zfill(4), string.capwords(name), string.capwords("winterstr"), "19", "60489", string.capwords(city), "01627466946", "12/08/1985", pay, "mhdaimi@hotmail.com",
                           "10/10/2011", "", gender]

    cur.execute(sql,parameters)
    conn.commit()