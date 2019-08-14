# -*- coding: cp1251 -*-

import config

import sqlite3
import os.path

def CheckDB():
    if not os.path.isfile(config.DBPath):
        CreateDB()

def CreateDB():
    con = sqlite3.connect(config.DBPath)
    cur = con.cursor()    

    # xls table
    cur.execute("CREATE TABLE XLSSettings(Id INT, NoI INT, Area TEXT, St_Opts INT, Byp_Opts TEXT, Ext_BP INT)")
    con.commit()

    cur.execute("INSERT INTO XLSSettings VALUES(1, 1, '', 0, '', 1)")
    con.commit()

    # fhx table
    cur.execute("CREATE TABLE FHXSettings(Id INT, XLSPath TEXT, Lang INT, BP_Name TEXT, BP_REF TEXT, DST INT, Decpt INT, Ext_B INT, DO_Opts TEXT, FB INT)")
    con.commit()

    cur.execute("INSERT INTO FHXSettings VALUES(1, '', 0, '', '', 0, 0, 0, '', 0)")
    con.commit()


def GetXLSData():
    con = sqlite3.connect(config.DBPath)
    cur = con.cursor()
    
    cur.execute("SELECT * FROM XLSSettings")
    
    row = cur.fetchone()
    con.close()
    
    return row

def GetFHXData():
    con = sqlite3.connect(config.DBPath)
    cur = con.cursor()
    
    cur.execute("SELECT * FROM FHXSettings")
    
    row = cur.fetchone()
    con.close()
    
    return row

def UpdateXLSData(NoI, Area, St_Opts, Byp_Opts, Ext_BP):
    con = sqlite3.connect(config.DBPath)
    cur = con.cursor()
    
    cur.execute("UPDATE XLSSettings SET NoI=?, Area=?, St_Opts=?, Byp_Opts=?, Ext_BP=? WHERE Id=?", (NoI, Area, St_Opts, Byp_Opts, Ext_BP, 1)) 
    
    con.commit()
    con.close()

def UpdateFHXData(XLSPath, Lang, BP_Name, BP_REF, DST, Decpt, Ext_B, DO_Opts, FB):
    con = sqlite3.connect(config.DBPath)
    cur = con.cursor()
    
    cur.execute("UPDATE FHXSettings SET XLSPath=?, Lang=?, BP_Name=?, BP_REF=?, DST=?, Decpt=?, Ext_B=?, DO_Opts=?, FB=? WHERE Id=?", (XLSPath, Lang, BP_Name, BP_REF, DST, Decpt, Ext_B, DO_Opts, FB, 1)) 
    
    con.commit()
    con.close()
