from openpyxl import Workbook
import mysql.connector
from openpyxl.styles import Font
from itertools import *
from mysql.connector.cursor import MySQLCursor
import pandas as pd
from pandas import ExcelWriter
import driver
import xlsxwriter

def Associatedcounts():
    conn = mysql.connector.connect(user=driver.user, password=driver.password, host=driver.host,database=driver.databasename)

# for busServiceId,COUNT(busServiceId )
    #part1
    mycursor = MySQLCursor(conn)
    mycursor.execute('SELECT busServiceId,COUNT(busServiceId ) FROM available_trips GROUP BY busServiceId')
    tablename = mycursor.fetchall()

    returndict={}
    tablename = dict(tablename)
    #print(tablename)

    returndict=tablename
   # print(returndict)
    var1 = tablename.keys()
    listvarkeys = list(var1)
   # print(listvarkeys)

    var2 = tablename.values()
    listvarvalues = list(var2)

    if (listvarkeys[0] == ''):
        listvarkeys[0]='Blank'

    tablename = {}
    for key in listvarkeys:
        for value in listvarvalues:
            tablename[key] = value
            listvarvalues.remove(value)
            break
    #FOR RETURN VALUE

    pd.DataFrame(tablename.items())
    svaldf = pd.DataFrame(tablename.items(), columns=['busServiceId', 'TOTAL COUNT(busServiceId) IN TABLE'])



  #part2

    mycursor = MySQLCursor(conn)
    mycursor.execute('SELECT busServiceId,COUNT(busServiceId ) FROM available_trips GROUP BY busServiceId ORDER BY count(busServiceId) DESC ')
    tablename = mycursor.fetchall()

    tablename = dict(tablename)
    var1 = tablename.keys()
    listvarkeys = list(var1)
    #print(listvarkeys)

    var2 = tablename.values()
    listvarvalues = list(var2)

    if (listvarkeys[0] == ''):
        listvarkeys[0] = 'Blank'

    tablename = {}
    for key in listvarkeys:
        for value in listvarvalues:
            tablename[key] = value
            listvarvalues.remove(value)
            break



    pd.DataFrame(tablename.items())
    svaldff = pd.DataFrame(tablename.items(), columns=['busServiceId', 'TOTAL COUNT(busServiceId) IN TABLE DESC'])

    writer_object = pd.ExcelWriter(driver.desktop+'Associated_counts\\ List_BusServiceID.xlsx')
    svaldf.to_excel(writer_object, startcol=0, startrow=1,sheet_name='Sheet1')
    svaldff.to_excel(writer_object, startcol=4, startrow=1,sheet_name='Sheet1')

    workbook_object = writer_object.book
    worksheet_object = writer_object.sheets['Sheet1']

    worksheet_object.set_column('B:C', 20)
    worksheet_object.set_column('F:G', 30)

    worksheet_object.write('B1', "Total rows of Busserviceid ")
    worksheet_object.write('C1', len(svaldff.axes[0]))

    writer_object.save()


    return returndict



def Associatedcounts2():
    conn = mysql.connector.connect(user=driver.user, password=driver.password, host=driver.host,database=driver.databasename)

# for busServiceId,COUNT(busServiceId )
    #part1
    mycursor = MySQLCursor(conn)
    mycursor.execute('SELECT tatkalTime,COUNT(tatkalTime ) FROM available_trips GROUP BY tatkalTime')
    tablename = mycursor.fetchall()

    tablename = dict(tablename)
    var1 = tablename.keys()
    listvarkeys = list(var1)
    #print(listvarkeys)

    var2 = tablename.values()
    listvarvalues = list(var2)

    if (listvarkeys[0] == ''):
        listvarkeys[0]='Blank'

    tablename = {}
    for key in listvarkeys:
        for value in listvarvalues:
            tablename[key] = value
            listvarvalues.remove(value)
            break
    #print("Resultant dictionary is : " + str(tablename))


    pd.DataFrame(tablename.items())
    svaldf = pd.DataFrame(tablename.items(), columns=['tatkalTime', 'TOTAL COUNT(tatkalTime) IN TABLE'])

  #part2

    mycursor = MySQLCursor(conn)
    mycursor.execute('SELECT tatkalTime,COUNT(tatkalTime ) FROM available_trips GROUP BY tatkalTime ORDER BY count(tatkalTime) DESC ')
    tablename = mycursor.fetchall()

    tablename = dict(tablename)
    var1 = tablename.keys()
    listvarkeys = list(var1)
    #print(listvarkeys)

    var2 = tablename.values()
    listvarvalues = list(var2)

    if (listvarkeys[0] == ''):
        listvarkeys[0] = 'Blank'

    tablename = {}
    for key in listvarkeys:
        for value in listvarvalues:
            tablename[key] = value
            listvarvalues.remove(value)
            break
    # print("Resultant dictionary is : " + str(tablename))


    #print("Resultant dictionary is : " + str(tablename))


    pd.DataFrame(tablename.items())
    svaldff = pd.DataFrame(tablename.items(), columns=['tatkalTime', 'TOTAL COUNT(tatkalTime) IN TABLE DESC'])

    writer_object = pd.ExcelWriter(driver.desktop+'Associated_counts\\ List_TatkalTiming.xlsx')
    svaldf.to_excel(writer_object, startcol=0, startrow=1,sheet_name='Sheet1')
    svaldff.to_excel(writer_object, startcol=4, startrow=1,sheet_name='Sheet1')

    workbook_object = writer_object.book
    worksheet_object = writer_object.sheets['Sheet1']

    worksheet_object.set_column('B:C', 20)
    worksheet_object.set_column('F:G', 30)

    worksheet_object.write('B1', "Total rows of tatkaltime ")
    worksheet_object.write('C1', len(svaldff.axes[0]))




    writer_object.save()

def Associatedcounts3():
    conn = mysql.connector.connect(user=driver.user, password=driver.password, host=driver.host,database=driver.databasename)

# for busServiceId,COUNT(busServiceId )
    #part1
    mycursor = MySQLCursor(conn)
    mycursor.execute('SELECT routeId,COUNT(routeId ) FROM available_trips GROUP BY routeId')
    tablename = mycursor.fetchall()

    tablename = dict(tablename)
    var1 = tablename.keys()
    listvarkeys = list(var1)
    #print(listvarkeys)

    var2 = tablename.values()
    listvarvalues = list(var2)

    if (listvarkeys[0] == ''):
        listvarkeys[0]='Blank'

    tablename = {}
    for key in listvarkeys:
        for value in listvarvalues:
            tablename[key] = value
            listvarvalues.remove(value)
            break
    #print("Resultant dictionary is : " + str(tablename))


    pd.DataFrame(tablename.items())
    svaldf = pd.DataFrame(tablename.items(), columns=['routeId', 'TOTAL COUNT(routeId) IN TABLE'])

  #part2

    mycursor = MySQLCursor(conn)
    mycursor.execute('SELECT routeId,COUNT(routeId ) FROM available_trips GROUP BY routeId ORDER BY count(routeId) DESC ')
    tablename = mycursor.fetchall()

    tablename = dict(tablename)
    var1 = tablename.keys()
    listvarkeys = list(var1)
    #print(listvarkeys)

    var2 = tablename.values()
    listvarvalues = list(var2)

    if (listvarkeys[0] == ''):
        listvarkeys[0] = 'Blank'

    tablename = {}
    for key in listvarkeys:
        for value in listvarvalues:
            tablename[key] = value
            listvarvalues.remove(value)
            break
    # print("Resultant dictionary is : " + str(tablename))


    #print("Resultant dictionary is : " + str(tablename))


    pd.DataFrame(tablename.items())
    svaldff = pd.DataFrame(tablename.items(), columns=['routeId', 'TOTAL COUNT(routeId) IN TABLE DESC'])

    writer_object = pd.ExcelWriter(driver.desktop+'Associated_counts\\List_routeId.xlsx')
    svaldf.to_excel(writer_object, startcol=0, startrow=1,sheet_name='Sheet1')
    svaldff.to_excel(writer_object, startcol=4, startrow=1,sheet_name='Sheet1')

    workbook_object = writer_object.book
    worksheet_object = writer_object.sheets['Sheet1']

    worksheet_object.set_column('B:C', 20)
    worksheet_object.set_column('F:G', 30)

    worksheet_object.write('B1', "Total rows of routeId ")
    worksheet_object.write('C1', len(svaldff.axes[0]))




    writer_object.save()