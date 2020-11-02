import sys
import os
import pyodbc
import xlsxwriter
from PyQt5 import QtCore, QtGui, QtWidgets
from test import Ui_MainWindow
app = QtWidgets.QApplication(sys.argv)
MainWindow = QtWidgets.QMainWindow()
ui = Ui_MainWindow()
def print_BD():
    conn = pyodbc.connect('DSN=PostgreSQL30tel')
    cursor = conn.cursor()
    query = cursor.execute("""
    select u_id,fam.fam,name.name,otc.otc,street.street,bldn,b_corp,appr,tel
    from main,fam,name,otc,street
    where main.fam=fam.f_id
    and main.name=name.n_id
    and main.otc=otc.o_id
    and main.street=street.s_id
    """)
    workbook = xlsxwriter.Workbook('BD.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write('A1','Номер')
    worksheet.write('B1','Фамилия')
    worksheet.write("C1",'Имя')
    worksheet.write("D1","Отчество")
    worksheet.write('E1',"Улица")
    worksheet.write('F1',"Дом")
    worksheet.write('G1',"Корпус")
    worksheet.write('H1',"Квартира")
    worksheet.write('I1',"Телефон")
    a = via(query)
    j = 1
    for b in a:
        c = ''.join(str(b))
        lst = c.replace("'"," ").split()
        for v in range (1,13,2):
            worksheet.write(j,v//2+1,lst[v].replace(","," "))
        worksheet.write(j,7,lst[13].replace(","," "))
        worksheet.write(j,8,lst[14].replace(","," "))
        worksheet.write(j,0,lst[0].replace("("," ").replace(",",' '))
        j+=1
    workbook.close() 
    os.startfile('BD.xlsx')
    conn.close()
def via (query):
    return (query.fetchall())
def new_data (fam,name,otc,strt,bldn,corp,app,tel):
    conn = pyodbc.connect('DSN=PostgreSQL30tel')
    cursor = conn.cursor()
    f_id = plus_fam(fam)
    n_id = plus_name(name)
    o_id = plus_otc(otc)
    s_id = plus_strt(strt)
    bldn_a = int(bldn)
    corp_a = int (corp)
    app_a = int (app)
    query = cursor.execute("""insert into main values (default,?,?,?,?,?,?,?,?)""",(f_id,n_id,o_id,s_id,bldn,corp,app_a,tel))
    conn.commit()
    conn.close()
def plus_fam(fam):
    conn = pyodbc.connect('DSN=PostgreSQL30tel')
    cursor = conn.cursor()
    query = cursor.execute("""select * from fam""")
    a = query.fetchall()
    f_id, m = check (a, fam)
    if f_id != 0:
        return (f_id)
    query = cursor.execute("""insert into fam values (default,?)""",(fam))
    conn.commit()
    query = cursor.execute("""select * from fam""")
    a = query.fetchall()
    f_id, m = check (a, fam)
    if f_id != 0:
        return (f_id)
    else :
        return ('False')
    conn.close()
def check (a, data):
    m = 0
    id = 0
    for b in a:
        c = ''.join(str(b))
        lst = c.replace("("," ").replace(")"," ").replace("'"," ").replace(","," ").split()
        if lst[1] == data:
            id = int(lst[0])
            m =+1
    return (id, m)
def plus_name(name):
    conn = pyodbc.connect('DSN=PostgreSQL30tel')
    cursor = conn.cursor()
    query = cursor.execute("""select * from name""")
    a = query.fetchall()
    f_id, m = check (a, name)
    if f_id != 0:
        return (f_id)
    query = cursor.execute("""insert into name values (default,?)""",(name))
    conn.commit()
    query = cursor.execute("""select * from name""")
    a = query.fetchall()
    f_id, m = check (a, name)
    if f_id != 0:
        return (f_id)
    else :
        return ('False')
    conn.close()
def plus_otc(otc):
    conn = pyodbc.connect('DSN=PostgreSQL30tel')
    cursor = conn.cursor()
    query = cursor.execute("""select * from otc""")
    a = query.fetchall()
    f_id, m = check (a, otc)
    if f_id != 0:
        return (f_id)
    query = cursor.execute("""insert into otc values (default,?)""",(otc))
    conn.commit()
    query = cursor.execute("""select * from otc""")
    a = query.fetchall()
    f_id, m = check (a, otc)
    if f_id != 0:
        return (f_id)
    else :
        return ('False')
    conn.close()
def plus_strt(strt):
    conn = pyodbc.connect('DSN=PostgreSQL30tel')
    cursor = conn.cursor()
    query = cursor.execute("""select * from street""")
    a = query.fetchall()
    f_id, m = check (a, strt)
    if f_id != 0:
        return (f_id)
    query = cursor.execute("""insert into street values (default,?)""",(strt))
    conn.commit()
    query = cursor.execute("""select * from street""")
    a = query.fetchall()
    f_id, m = check (a, strt)
    if f_id != 0:
        return (f_id)
    else :
        return ('False')
    conn.close()
