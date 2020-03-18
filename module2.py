from openpyxl import Workbook
import mysql.connector
from openpyxl.styles import Font
from itertools import *
import openpyxl
from mysql.connector.cursor import MySQLCursor
import pandas as pd
from pandas import ExcelWriter
import driver
import xlsxwriter
import numpy as np
from collections import Counter
from  more_itertools import unique_everseen


def Combinedssociatedcounts2():
    conn = mysql.connector.connect(user=driver.user, password=driver.password, host=driver.host,database=driver.databasename)

    dict1 = {}
    dict2 = {}
    dict3 = {}
    dict4 = {}

#for busservicid counts occurence
    mycursor = MySQLCursor(conn)
    mycursor.execute('SELECT busServiceId FROM available_trips ORDER by busServiceId ')
    myresult65 = mycursor.fetchall()  # sample12)avlWindowSeats   30 collect from table
    busserviceidoutput = [item for x in zip_longest(*myresult65) for item in x if item != -55]
    #print(busserviceidoutput1)
    for x in range(len(busserviceidoutput)):
        if busserviceidoutput[x] == '':
            busserviceidoutput[x] = "BLANK"
   #busserviceidoutput=list(unique_everseen(busserviceidoutput)) #removing duplicates

    #print(busserviceidoutput)

    c = Counter(busserviceidoutput)
   # print(c)                              #counting ocurrence
    finaldict = dict(c)
    dict1=finaldict
    #print(finaldict)

    pd.DataFrame(finaldict.items())
    busservicedf = pd.DataFrame(finaldict.items(), columns=['busServiceId', 'busServiceId_Unique_Count'])
    writer_object = pd.ExcelWriter(driver.desktop+'Combinedssociatedcounts2.xlsx')
    busservicedf.to_excel(writer_object, startcol=0, startrow=1, sheet_name='Sheet1')


    #worksheet_object.set_column('F:G', 30)


# for travels counts occurence
    mycursor = MySQLCursor(conn)
    mycursor.execute('SELECT travels FROM available_trips ORDER by travels ')
    myresult65 = mycursor.fetchall()  # sample12)avlWindowSeats   30 collect from table
    travelsoutput = [item for x in zip_longest(*myresult65) for item in x if item != -55]
    # print(busserviceidoutput1)
    for x in range(len(travelsoutput)):
        if travelsoutput[x] == '':
            travelsoutput[x] = "BLANK"
    # busserviceidoutput=list(unique_everseen(busserviceidoutput)) #removing duplicates

    c = Counter(travelsoutput)
    #print(c)  # counting ocurrence
    finaldict = dict(c)
    dict2 = finaldict

    pd.DataFrame(finaldict.items())
    travelsdf = pd.DataFrame(finaldict.items(), columns=['travels', 'travels_Unique_Count'])
    travelsdf.to_excel(writer_object, startcol=4, startrow=1, sheet_name='Sheet1')



# for  	routeId  counts occurence
    mycursor = MySQLCursor(conn)
    mycursor.execute('SELECT  routeId  FROM available_trips ORDER by  routeId  ')
    myresult65 = mycursor.fetchall()  # sample12)avlWindowSeats   30 collect from table
    routeIdoutput = [item for x in zip_longest(*myresult65) for item in x if item != -55]
    # print(busserviceidoutput1)
    for x in range(len(routeIdoutput)):
        if routeIdoutput[x] == '':
            routeIdoutput[x] = "BLANK"
    # busserviceidoutput=list(unique_everseen(busserviceidoutput)) #removing duplicates

    c = Counter(routeIdoutput)
    #print(c)  # counting ocurrence
    finaldict = dict(c)

    dict3 = finaldict

    pd.DataFrame(finaldict.items())
    routeIddf = pd.DataFrame(finaldict.items(), columns=[' routeId ', 'routeId _Unique_Count'])
    routeIddf.to_excel(writer_object, startcol=8, startrow=1, sheet_name='Sheet1')



# for  operator  counts occurence
    mycursor = MySQLCursor(conn)
    mycursor.execute('SELECT operator FROM available_trips ORDER by  operator  ')
    myresult65 = mycursor.fetchall()
    operatoroutput = [item for x in zip_longest(*myresult65) for item in x if item != -55]
    # print(busserviceidoutput1)
    for x in range(len(operatoroutput)):
        if operatoroutput[x] == '':
            operatoroutput[x] = "BLANK"
    # busserviceidoutput=list(unique_eveBlankrseen(busserviceidoutput)) #removing duplicates

    c = Counter(operatoroutput)
    #print(c)  # counting ocurrence
    finaldict = dict(c)
    dict4 = finaldict

    pd.DataFrame(finaldict.items())
    operatordf = pd.DataFrame(finaldict.items(), columns=[' operator ', 'operator _Unique_Count'])
    operatordf.to_excel(writer_object, startcol=12, startrow=1, sheet_name='Sheet1')

 # for  bustypeID  counts occurence
    mycursor = MySQLCursor(conn)
    mycursor.execute('SELECT busTypeId FROM available_trips ORDER by  busTypeId  ')
    myresult65 = mycursor.fetchall()
    busTypeIdoutput = [item for x in zip_longest(*myresult65) for item in x if item != -55]
    # print(busserviceidoutput1)
    for x in range(len(busTypeIdoutput)):
        if busTypeIdoutput[x] == '':
            busTypeIdoutput[x] = "BLANK"
    # busserviceidoutput=list(unique_everseen(busserviceidoutput)) #removing duplicates

    c = Counter(busTypeIdoutput)
    #print(c)  # counting ocurrence
    finaldict = dict(c)
    dict5 = finaldict

    pd.DataFrame(finaldict.items())
    busTypeIddf = pd.DataFrame(finaldict.items(), columns=[' 	busTypeId ', '	busTypeId _Unique_Count'])
    busTypeIddf.to_excel(writer_object, startcol=16, startrow=1, sheet_name='Sheet1')



    worksheet_object = writer_object.sheets['Sheet1']

    worksheet_object.set_column('A:C', 30)
    worksheet_object.set_column('E:G', 30)
    worksheet_object.set_column('I:K', 30)
    worksheet_object.set_column('M:O', 30)
    worksheet_object.set_column('Q:S', 30)


    worksheet_object.write('B1', "Total rows of Busserviceid ")
    worksheet_object.write('C1', len(busservicedf.axes[0]))

    worksheet_object.write('F1', "Total rows of travels")
    worksheet_object.write('G1', len(travelsdf.axes[0]))

    worksheet_object.write('J1', "Total rows of routeid ")
    worksheet_object.write('K1', len(routeIddf.axes[0]))

    worksheet_object.write('N1', "Total rows of operator ")
    worksheet_object.write('O1', len(operatordf.axes[0]))

    worksheet_object.write('R1', "Total rows of busTypeId ")
    worksheet_object.write('S1', len(busTypeIddf.axes[0]))

    writer_object.save()

    # print 4 attribs as per types sending dataframe to another function
    ###################for busserviceid

    s1 = list(dict1.keys())

    if (s1[0] == ''):
        s1[0] = 'BLANK'

    s2 = list(dict1.values())
    data = {'type': 'Busserviceid', 'id': s1, 'count': s2}
    dff1 = pd.DataFrame(data)

    ###############################for travels
    s1 = list(dict2.keys())

    if (s1[0] == ''):
        s1[0] = 'BLANK'

    s2 = list(dict2.values())
    data = {'type': 'travels', 'id': s1, 'count': s2}
    dff2 = pd.DataFrame(data)
    ######################## for routeId
    s1 = list(dict3.keys())
    # print(s1)
    for x in range(len(s1)):
        s1[x] = str(s1[x])
    # print(s1)

    if (s1[0] == ''):
        s1[0] = 'BLANK'

    s2 = list(dict3.values())
    data = {'type': 'routeId', 'id': s1, 'count': s2}
    dff3 = pd.DataFrame(data)
    ################for operator
    s1 = list(dict4.keys())
    if (s1[0] == ''):
        s1[0] = 'BLANK'

    s2 = list(dict4.values())
    data = {'type': 'Operator', 'id': s1, 'count': s2}
    dff4 = pd.DataFrame(data)


    ################for bustypeid

    s1 = list(dict5.keys())
    if (s1[0] == ''):
        s1[0] = 'BLANK'

    s2 = list(dict5.values())
    data = {'type': 'busTypeId', 'id': s1, 'count': s2}
    dff5 = pd.DataFrame(data)


    dff1 = dff1.append(dff2, ignore_index=True)
    dff1 = dff1.append(dff3, ignore_index=True)
    dff1 = dff1.append(dff4, ignore_index=True)
    dff1 = dff1.append(dff5, ignore_index=True)


    conn.close()
    # print(dff1)

    return dff1


# df2.iloc[0].

def dataframewrite(value):
    writer_object = pd.ExcelWriter(driver.desktop+'type_serialno_id_count.xlsx')
    value.to_excel(writer_object, startcol=0, startrow=0, sheet_name='Sheet1')

    workbook_object = writer_object.book
    worksheet_object = writer_object.sheets['Sheet1']

    worksheet_object.set_column('B:C', 30)

    writer_object.save()












      

    



  
  
