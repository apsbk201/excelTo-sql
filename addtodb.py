#!/usr/bin/python3
import openpyxl
import pandas as  pd
import sqlite3

db = 'table.db'
conn = sqlite3.connect(db)
cur = conn.cursor()


def createTable_employee():
    try:
        cur.execute("""
        CREATE TABLE IF NOT EXISTS employee (
        id INTEGER PRIMARY KEY,
        fname VARCHAR(255) NOT NULL,
        lname VARCHAR(255) NOT NULL,
        address VARCHAR(255) NOT NULL,
        email VARCHAR(255) NOT NULL,
        tel VARCHAR(255) NOT NULL,
        contact VARCHAR(255) NOT NULL,
        status VARCHAR(255) NOT NULL,
        position VARCHAR(255) NOT NULL,
        permistion VARCHAR(255) NOT NULL,
        dateIn VARCHAR(255) NOT NULL,
        dateOut VARCHAR(255) NOT NULL,
        uname VARCHAR(255) NOT NULL,
        password VARCHAR(255) NOT NULL,
        hint VARCHAR(255) NOT NULL
        )
        """)
        # print('created table employee')
    except Exception as e:
        print(e)

def createTable_vender():
    try:
        cur.execute("""
        CREATE TABLE IF NOT EXISTS vender (
        id INTEGER PRIMARY KEY,
        fname VARCHAR(255) NOT NULL,
        lname VARCHAR(255) NOT NULL,
        address VARCHAR(255) NOT NULL,
        tax VARCHAR(255) NOT NULL,
        email VARCHAR(255) NOT NULL,
        tel VARCHAR(255) NOT NULL,
        contact VARCHAR(255) NOT NULL,
        score VARCHAR(255) NOT NULL,
        comment VARCHAR(255) NOT NULL
        )
        """)
        # print('created table vender')
    except Exception as e:
        print(f'create vender {e}')


def createTable_customer():
    try:
        cur.execute("""
        CREATE TABLE IF NOT EXISTS customer (
        id INTEGER PRIMARY KEY,
        fname VARCHAR(255) NOT NULL,
        lname VARCHAR(255) NOT NULL,
        address VARCHAR(255) NOT NULL,
        tax VARCHAR(255) NOT NULL,
        email VARCHAR(255) NOT NULL,
        tel VARCHAR(255) NOT NULL,
        contact VARCHAR(255) NOT NULL,
        score VARCHAR(255) NOT NULL,
        comment VARCHAR(255) NOT NULL
        )
        """)
        # print('created table employee')
    except Exception as e:
        print(e)

def createTable_product():
    try:
        cur.execute("""
        CREATE TABLE IF NOT EXISTS product (
        id INTEGER PRIMARY KEY,
        model VARCHAR(255) NOT NULL,
        name VARCHAR(255) NOT NULL,
        serial VARCHAR(255) NOT NULL,
        unit VARCHAR(255) NOT NULL,
        mfg VARCHAR(255) NOT NULL,
        exp VARCHAR(255) NOT NULL,
        lot VARCHAR(255) NOT NULL,
        dateIn VARCHAR(255) NOT NULL,
        dateOut VARCHAR(255) NOT NULL,
        venderId VARCHAR(255) NOT NULL,
        venderSn VARCHAR(255) NOT NULL,
        price VARCHAR(255) NOT NULL,
        quantity INTEGER NOT NULL,
        out INTEGER NOT NULL,
        balance INTEGER NOT NULL,
        comment VARCHAR(255) NOT NULL
        )
        """)
        # print('created table employee')
    except Exception as e:
        print(e)

def createTable_invoce():
    try:
        cur.execute("""
        CREATE TABLE IF NOT EXISTS invoice (
        id INTEGER PRIMARY KEY,
        date VARCHAR(255) NOT NULL,
        invoiceNumber VARCHAR(255) NOT NULL,
        productId VARCHAR(255) NOT NULL,
        quantity VARCHAR(255) NOT NULL,
        unit VARCHAR(255) NOT NULL,
        price VARCHAR(255) NOT NULL,
        dateOut VARCHAR(255) NOT NULL,
        waranty VARCHAR(255) NOT NULL,
        service VARCHAR(255) NOT NULL,
        customerId VARCHAR(255) NOT NULL,
        saleId VARCHAR(255) NOT NULL,
        comment VARCHAR(255) NOT NULL
        )
        """)
        # print('created table employee')
    except Exception as e:
        print(e)

createTable_employee()
createTable_vender()
createTable_customer()
createTable_product()
createTable_invoce()
def insertTable_employee():
    try:
        emp = pd.read_excel('company.xlsx',sheet_name='employee')
        employee = emp.to_records(index=False)
        employee = list(employee)
        for df in employee:
            sql = 'INSERT INTO employee (fname,lname,address,email,tel,contact,status,position,permistion,dateIn,dateOut,uname,password,hint) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)'
            value = (str(df[0]),str(df[1]),str(df[2]), str(df[3]),str(df[4]),str(df[5]),str(df[6]),str(df[7]),str(df[8]),str(df[9]),str(df[10]),str(df[11]),str(df[12]),str(df[13]))
            cur.execute(sql,value)
        print('employee Done')
        conn.commit()
    except Exception as e:
        print(e)

def insertTable_vender():
    try:
        ven = pd.read_excel('company.xlsx',sheet_name='vender')
        vender = ven.to_records(index=False)
        vender = list(vender)
        for df in vender:
            sql = "INSERT INTO vender (fname,lname, address, tax, email,tel,contact,score,comment)VALUES (?,?,?,?,?,?,?,?,?)"
            value = (str(df[0]),str(df[1]),str(df[2]), str(df[3]),str(df[4]),str(df[5]),str(df[6]),str(df[7]),str(df[8]))
            cur.execute(sql,value)
        print('vender Done')
        conn.commit()
    except Exception as e:
        print(e)

def insertTable_customer():
    try:
        cus = pd.read_excel('company.xlsx',sheet_name='customer')
        customer = cus.to_records(index=False)
        customer = list(customer)
        for df in customer:

            sql = "INSERT INTO customer (fname,lname, address, tax, email,tel,contact,score,comment)VALUES (?,?,?,?,?,?,?,?,?)"
            value = (str(df[0]),str(df[1]),str(df[2]), str(df[3]),str(df[4]),str(df[5]),str(df[6]),str(df[7]),str(df[8]))
            cur.execute(sql,value)
        print('customer Done')
        conn.commit()
    except Exception as e:
        print(e)

def insertTable_product():
    try:
        pro = pd.read_excel('company.xlsx',sheet_name='product')
        product = pro.to_records(index=False)
        product = list(product)
        for df in product:
            sql = "INSERT INTO product (model,name,serial,unit,mfg,exp,lot,dateIn,dateOut,venderId,venderSn,price,quantity,out,balance,comment)VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
            value = (str(df[0]),str(df[1]),str(df[2]), str(df[3]),str(df[4]),str(df[5]),str(df[6]),str(df[7]),str(df[8]),str(df[9]),str(df[10]),str(df[11]),str(df[12]),str(df[13]),str(df[14]),str(df[15]))
            cur.execute(sql,value)
        print('product Done')
        conn.commit()
    except Exception as e:
        print(e)
def insertTable_invoice():
    try:
        inv = pd.read_excel('company.xlsx',sheet_name='invoice')
        invoice = inv.to_records(index=False)
        invoice = list(invoice)
        for df in invoice:
            sql = "INSERT INTO invoice (date,invoiceNumber,productId,quantity,unit,price,dateOut,waranty,service,customerId,saleId,comment)VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"
            value = (str(df[0]),str(df[1]),str(df[2]), str(df[3]),str(df[4]),str(df[5]),str(df[6]),str(df[7]),str(df[8]),str(df[9]),str(df[10]),str(df[11]))
            cur.execute(sql,value)
        print('invoice Done')
        conn.commit()
    except Exception as e:
        print(e)

insertTable_employee()
insertTable_vender()
insertTable_customer()
insertTable_product()
insertTable_invoice()
