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
from more_itertools import unique_everseen
import xlrd
import numpy as np


def refbusserviceidcount():
    conn = mysql.connector.connect(user=driver.user, password=driver.password, host=driver.host,
                                   database=driver.databasename)
    mycursor = MySQLCursor(conn)
    mycursor.execute(
        'SELECT  DISTINCT(busServiceId),travels,routeId,operator,busTypeId FROM available_trips ORDER BY busServiceId ASC')
    refbusserviceid = mycursor.fetchall()
    # dfrefbusserviceid
    # A column
    dfrefbusserviceid = pd.DataFrame(refbusserviceid)  # main dataframe
    dfrefbusserviceid.columns = ['busservicid', 'travels', 'routeid',
                                 'operator', 'bustypeid', ]
    dfrefbusserviceid = dfrefbusserviceid.drop_duplicates(subset='busservicid', keep='first', inplace=False)
    dfrefbusserviceid.replace('', 'BLANK', inplace=True)
    # print(dfrefbusserviceid)

    writer_object = pd.ExcelWriter(driver.desktop + '1_refbusserviceid.xlsx')

    # dfrefbusserviceid['busservicid'].to_excel(writer_object, startcol=0, startrow=1, sheet_name='Sheet1',index= False)

    # B column

    tempdf = driver.dataframe1[(driver.dataframe1.type.isin(['Busserviceid']))]
    tempdf = tempdf.drop(['type'], axis=1)

    tempdfdict = tempdf.set_index('id')['count'].to_dict()
    # print(tempdfdict)
    finaldf = pd.DataFrame()
    listbcol = dfrefbusserviceid['busservicid'].tolist()
    # print(listbcol)
    finaldf = pd.DataFrame({
        'buserviceid': listbcol,
        'Unique_count_buserviceid': listbcol})

    finaldf['Unique_count_buserviceid'].replace(tempdfdict)

    # print(finaldf)

    finaldf.to_excel(writer_object, startcol=0, startrow=2, sheet_name='Sheet1', index=False, header=None)
    # C Column

    #  dfrefbusserviceid['travels'].to_excel(writer_object, startcol=2, startrow=2, sheet_name='Sheet1',index= False,header=None)

    # D column

    travelslist = dfrefbusserviceid['travels'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf = driver.dataframe1[(driver.dataframe1.type.isin(['travels']))]
    tempdf = tempdf.drop(['type'], axis=1)
    tempdfdict = tempdf.set_index('id')['count'].to_dict()

    # print(tempdfdict)

    finaldfdcol = pd.DataFrame(travelslist, columns=['travels '])
    finaldfdcol['travels_count'] = travelslist

    finaldfdcol = finaldfdcol.replace({"travels_count": tempdfdict})
    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldfdcol.to_excel(writer_object, startcol=2, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # print(finaldf)
    # E & F column routeid

    routeidlist = dfrefbusserviceid['routeid'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf1 = driver.dataframe1[(driver.dataframe1.type.isin(['routeId']))]
    tempdf1 = tempdf1.drop(['type'], axis=1)
    tempdf1dict = tempdf1.set_index('id')['count'].to_dict()

    # print(tempdf1dict)

    routeidlist = list(map(str, routeidlist))  # casting list to str

    finaldfedcol = pd.DataFrame(routeidlist, columns=['routeId '])
    # print(finaldfdcol)
    finaldfedcol['routeId_count'] = routeidlist
    # print(finaldfedcol)

    finaldfedcol = finaldfedcol.replace({"routeId_count": tempdf1dict})

    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldfedcol.to_excel(writer_object, startcol=4, startrow=2, sheet_name='Sheet1', index=False, header=None)

# G and H column

    operatorlist = dfrefbusserviceid['operator'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf2 = driver.dataframe1[(driver.dataframe1.type.isin(['Operator']))]
    tempdf2 = tempdf2.drop(['type'], axis=1)
    tempdf2dict = tempdf2.set_index('id')['count'].to_dict()

   # print(tempdf2dict)

    finaldf2dcol = pd.DataFrame(operatorlist, columns=['operator '])
    finaldf2dcol['operator_count'] = operatorlist



    #print(finaldfdcol)

    finaldf2dcol = finaldf2dcol.replace(to_replace="operator_count",value=tempdf2dict)
    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldf2dcol.to_excel(writer_object, startcol=6, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # I and J column

    busTypeIdlist = dfrefbusserviceid['bustypeid'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf3 = driver.dataframe1[(driver.dataframe1.type.isin(['busTypeId']))]
    tempdf3 = tempdf3.drop(['type'], axis=1)
    tempdf3dict = tempdf3.set_index('id')['count'].to_dict()

    # print(tempdfdict)

    finaldf3dcol = pd.DataFrame(busTypeIdlist, columns=['busTypeId '])
    finaldf3dcol['busTypeId_count'] = busTypeIdlist

    finaldf3dcol['busTypeId_count'].replace(tempdf3dict)

    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldf3dcol.to_excel(writer_object, startcol=8, startrow=2, sheet_name='Sheet1', index=False, header=None)

    workbook_object = writer_object.book
    worksheet_object = writer_object.sheets['Sheet1']
    worksheet_object.set_column('A:B', 30)
    worksheet_object.set_column('C:D', 30)
    worksheet_object.set_column('E:F', 30)
    worksheet_object.set_column('G:H', 30)
    worksheet_object.set_column('I:J', 30)

    worksheet_object.write('A2', "Buserviceid ")

    worksheet_object.write('B2', "Unique_count_Buserviceid ")
    worksheet_object.write('C2', "travels ")
    worksheet_object.write('D2', "count_Travels ")
    worksheet_object.write('E2', " routeId")
    worksheet_object.write('F2', "count_routeId ")
    worksheet_object.write('G2', " operator")
    worksheet_object.write('H2', "count_operator ")
    worksheet_object.write('I2', " busTypeId")
    worksheet_object.write('J2', "count_busTypeId ")

    worksheet_object.write('A1', "Total_rows_all _col:" + str(dfrefbusserviceid.shape[0]))

    writer_object.save()


def reftravelscount():
    conn = mysql.connector.connect(user=driver.user, password=driver.password, host=driver.host,
                                   database=driver.databasename)
    mycursor = MySQLCursor(conn)
    mycursor.execute(
        'SELECT  DISTINCT(travels),busServiceId,routeId,operator,busTypeId FROM available_trips ORDER BY travels ASC')
    reftravels = mycursor.fetchall()
    # A column
    dfreftravels = pd.DataFrame(reftravels)  # main dataframe
    dfreftravels.columns = ['travels', 'busservicid', 'routeid',
                            'operator', 'bustypeid', ]
    dfreftravels = dfreftravels.drop_duplicates(subset='travels', keep='first', inplace=False)

    dfreftravels.replace('', 'BLANK', inplace=True)

    writer_object = pd.ExcelWriter(driver.desktop + '2_reftravels.xlsx')

    dfreftravels['travels'].to_excel(writer_object, startcol=0, startrow=1, sheet_name='Sheet1', index=False)

    # B column
    tempdf = driver.dataframe1[(driver.dataframe1.type.isin(['travels']))]
    tempdf = tempdf.drop(['type'], axis=1)

    tempdfdict = tempdf.set_index('id')['count'].to_dict()
    # print(tempdfdict)
    finaldf = pd.DataFrame()
    listbcol = dfreftravels['travels'].tolist()

    finaldf = pd.DataFrame({
        'travels': listbcol,
        'Unique_count_travels': listbcol})

    finaldf['Unique_count_travels'].replace(tempdfdict)

    finaldf.to_excel(writer_object, startcol=0, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # C and D column

    travelslist = dfreftravels['busservicid'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf = driver.dataframe1[(driver.dataframe1.type.isin(['Busserviceid']))]
    tempdf = tempdf.drop(['type'], axis=1)
    tempdfdict = tempdf.set_index('id')['count'].to_dict()

    # print(tempdfdict)

    finaldfdcol = pd.DataFrame(travelslist, columns=['busservicid '])
    finaldfdcol['busservicid_count'] = travelslist

    finaldfdcol = finaldfdcol.replace({"busservicid_count": tempdfdict})
    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldfdcol.to_excel(writer_object, startcol=2, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # print(finaldf)
    # E & F column routeid

    routeidlist = dfreftravels['routeid'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf1 = driver.dataframe1[(driver.dataframe1.type.isin(['routeId']))]
    tempdf1 = tempdf1.drop(['type'], axis=1)
    tempdf1dict = tempdf1.set_index('id')['count'].to_dict()

    # print(tempdf1dict)

    routeidlist = list(map(str, routeidlist))  # casting list to str

    finaldfedcol = pd.DataFrame(routeidlist, columns=['routeId '])
    # print(finaldfdcol)
    finaldfedcol['routeId_count'] = routeidlist
    # print(finaldfedcol)

    finaldfedcol = finaldfedcol.replace({"routeId_count": tempdf1dict})

    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldfedcol.to_excel(writer_object, startcol=4, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # G and H column

    operatorlist = dfreftravels['operator'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf2 = driver.dataframe1[(driver.dataframe1.type.isin(['Operator']))]
    tempdf2 = tempdf2.drop(['type'], axis=1)
    tempdf2dict = tempdf2.set_index('id')['count'].to_dict()

    # print(tempdfdict)

    finaldf2dcol = pd.DataFrame(operatorlist, columns=['operator '])
    finaldf2dcol['operator_count'] = operatorlist

    finaldf2dcol = finaldf2dcol.replace(to_replace="operator_count",value=tempdf2dict)


  #  finaldf2dcol = finaldf2dcol.replace({"operator_count": tempdf2dict})
    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldf2dcol.to_excel(writer_object, startcol=6, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # I and J column

    busTypeIdlist = dfreftravels['bustypeid'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf3 = driver.dataframe1[(driver.dataframe1.type.isin(['busTypeId']))]
    tempdf3 = tempdf3.drop(['type'], axis=1)
    tempdf3dict = tempdf3.set_index('id')['count'].to_dict()

    # print(tempdfdict)

    finaldf3dcol = pd.DataFrame(busTypeIdlist, columns=['busTypeId '])
    finaldf3dcol['busTypeId_count'] = busTypeIdlist

    finaldf3dcol['busTypeId_count'].replace(tempdf3dict)

    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldf3dcol.to_excel(writer_object, startcol=8, startrow=2, sheet_name='Sheet1', index=False, header=None)

    workbook_object = writer_object.book
    worksheet_object = writer_object.sheets['Sheet1']
    worksheet_object.set_column('A:B', 30)
    worksheet_object.set_column('C:D', 30)
    worksheet_object.set_column('E:F', 30)
    worksheet_object.set_column('G:H', 30)
    worksheet_object.set_column('I:J', 30)

    worksheet_object.write('A2', "travels ")

    worksheet_object.write('B2', "Unique_count_travels ")
    worksheet_object.write('C2', "busserviceid ")
    worksheet_object.write('D2', "count_busserviceid ")
    worksheet_object.write('E2', " routeId")
    worksheet_object.write('F2', "count_routeId ")
    worksheet_object.write('G2', " operator")
    worksheet_object.write('H2', "count_operator ")
    worksheet_object.write('I2', " busTypeId")
    worksheet_object.write('J2', "count_busTypeId ")

    worksheet_object.write('A1', "Total_rows_all _col:" + str(dfreftravels.shape[0]))

    writer_object.save()


def refrouteIDcount():
    conn = mysql.connector.connect(user=driver.user, password=driver.password, host=driver.host,
                                   database=driver.databasename)
    mycursor = MySQLCursor(conn)
    mycursor.execute(
        'SELECT DISTINCT(routeId) ,busServiceId,travels,operator,busTypeId FROM available_trips ORDER BY routeId ASC')
    refrouteID = mycursor.fetchall()
    # dfrefbusserviceid
    # A column
    dfrefrouteID = pd.DataFrame(refrouteID)  # main dataframe
    dfrefrouteID.columns = ['routeid', 'busservicid', 'travels',
                            'operator', 'bustypeid', ]
    dfrefrouteID = dfrefrouteID.drop_duplicates(subset='routeid', keep='first', inplace=False)

    dfrefrouteID.replace('', 'BLANK', inplace=True)

    writer_object = pd.ExcelWriter(driver.desktop + '3_refrouteID.xlsx')

    dfrefrouteID['routeid'].to_excel(writer_object, startcol=0, startrow=1, sheet_name='Sheet1', index=False)

# B column

    tempdf = driver.dataframe1[(driver.dataframe1.type.isin(['routeId']))]
    tempdf = tempdf.drop(['type'], axis=1)

    tempdfdict = tempdf.set_index('id')['count'].to_dict()

    # print(tempdfdict)
    finaldf = pd.DataFrame()
    listbcol = dfrefrouteID['routeid'].tolist()
    listbcol = list(map(str, listbcol))  # casting list to str

    finaldf = pd.DataFrame(listbcol, columns=['routeid '])
    finaldf['Unique_count_routeid'] = listbcol

    # print(finaldfdcol)

    finaldf = finaldf.replace(to_replace="Unique_count_routeid", value=tempdfdict)

   # finaldf = pd.DataFrame({
       # 'routeid': listbcol,
      #  'Unique_count_routeid': listbcol})

   # finaldf['Unique_count_routeid'].replace(tempdfdict)

    # print(finaldf)

    finaldf.to_excel(writer_object, startcol=0, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # C and D column

    busservicidlist = dfrefrouteID['busservicid'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf = driver.dataframe1[(driver.dataframe1.type.isin(['Busserviceid']))]
    tempdf = tempdf.drop(['type'], axis=1)
    tempdfdict = tempdf.set_index('id')['count'].to_dict()

    # print(tempdfdict)

    finaldfdcol = pd.DataFrame(busservicidlist, columns=['busservicid '])
    finaldfdcol['busservicid_count'] = busservicidlist

    finaldfdcol = finaldfdcol.replace({"busservicid_count": tempdfdict})
    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldfdcol.to_excel(writer_object, startcol=2, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # print(finaldf)

    # E & F column travels

    travelslist = dfrefrouteID['travels'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf1 = driver.dataframe1[(driver.dataframe1.type.isin(['travels']))]
    tempdf1 = tempdf1.drop(['type'], axis=1)
    tempdf1dict = tempdf1.set_index('id')['count'].to_dict()

    # print(tempdfdict)

    finaldf1dcol = pd.DataFrame(travelslist, columns=['travels '])
    finaldf1dcol['travels_count'] = travelslist

    finaldf1dcol = finaldf1dcol.replace({"travels_count": tempdf1dict})
    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldf1dcol.to_excel(writer_object, startcol=4, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # G and H column

    operatorlist = dfrefrouteID['operator'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf2 = driver.dataframe1[(driver.dataframe1.type.isin(['Operator']))]
    tempdf2 = tempdf2.drop(['type'], axis=1)
    tempdf2dict = tempdf2.set_index('id')['count'].to_dict()

    # print(tempdfdict)

    finaldf2dcol = pd.DataFrame(operatorlist, columns=['operator '])
    finaldf2dcol['operator_count'] = operatorlist

    finaldf2dcol = finaldf2dcol.replace(to_replace="operator_count",value=tempdf2dict)


   # finaldf2dcol = finaldf2dcol.replace({"operator_count": tempdf2dict})
    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldf2dcol.to_excel(writer_object, startcol=6, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # I and J column

    busTypeIdlist = dfrefrouteID['bustypeid'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf3 = driver.dataframe1[(driver.dataframe1.type.isin(['busTypeId']))]
    tempdf3 = tempdf3.drop(['type'], axis=1)
    tempdf3dict = tempdf3.set_index('id')['count'].to_dict()

    # print(tempdfdict)

    finaldf3dcol = pd.DataFrame(busTypeIdlist, columns=['busTypeId '])
    finaldf3dcol['busTypeId_count'] = busTypeIdlist

    finaldf3dcol['busTypeId_count'].replace(tempdf3dict)

    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldf3dcol.to_excel(writer_object, startcol=8, startrow=2, sheet_name='Sheet1', index=False, header=None)

    workbook_object = writer_object.book
    worksheet_object = writer_object.sheets['Sheet1']
    worksheet_object.set_column('A:B', 30)
    worksheet_object.set_column('C:D', 30)
    worksheet_object.set_column('E:F', 30)
    worksheet_object.set_column('G:H', 30)
    worksheet_object.set_column('I:J', 30)

    worksheet_object.write('A2', "routeId ")

    worksheet_object.write('B2', "Unique_count_routeId ")
    worksheet_object.write('C2', "busServiceId ")
    worksheet_object.write('D2', "count_busServiceId ")
    worksheet_object.write('E2', " travels")
    worksheet_object.write('F2', "count_travels ")
    worksheet_object.write('G2', " operator")
    worksheet_object.write('H2', "count_operator ")
    worksheet_object.write('I2', " busTypeId")
    worksheet_object.write('J2', "count_busTypeId ")

    worksheet_object.write('A1', "Total_rows_all _col:" + str(dfrefrouteID.shape[0]))

    writer_object.save()


def refoperatorcount():
    conn = mysql.connector.connect(user=driver.user, password=driver.password, host=driver.host,
                                   database=driver.databasename)
    mycursor = MySQLCursor(conn)
    mycursor.execute(
        'SELECT DISTINCT(operator),busServiceId,routeId,travels,busTypeId FROM available_trips ORDER BY operator ASC')
    refoperator = mycursor.fetchall()
    # dfrefbusserviceid
    # A column
    dfrefoperator = pd.DataFrame(refoperator)  # main dataframe
    dfrefoperator.columns = ['operator', 'busservicid', 'routeid', 'travels',
                             'bustypeid', ]

    dfrefoperator = dfrefoperator.drop_duplicates(subset='operator', keep='first', inplace=False)

    dfrefoperator.replace('', 'BLANK', inplace=True)

    writer_object = pd.ExcelWriter(driver.desktop + '4_refoperator.xlsx')

    dfrefoperator['operator'].to_excel(writer_object, startcol=0, startrow=1, sheet_name='Sheet1', index=False)

    # B column

    tempdf = driver.dataframe1[(driver.dataframe1.type.isin(['Operator']))]
    tempdf = tempdf.drop(['type'], axis=1)

    tempdfdict = tempdf.set_index('id')['count'].to_dict()
    # print(tempdfdict)
    finaldf = pd.DataFrame()
    listbcol = dfrefoperator['operator'].tolist()

    finaldf = pd.DataFrame({
        'operator': listbcol,
        'Unique_count_operator': listbcol})

    finaldf['Unique_count_operator'].replace(tempdfdict)

    finaldf.to_excel(writer_object, startcol=0, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # C and D column

    busServiceIdlist = dfrefoperator['busservicid'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf = driver.dataframe1[(driver.dataframe1.type.isin(['Busserviceid']))]
    tempdf = tempdf.drop(['type'], axis=1)
    tempdfdict = tempdf.set_index('id')['count'].to_dict()

    # print(tempdfdict)

    finaldfdcol = pd.DataFrame(busServiceIdlist, columns=['busservicid '])
    finaldfdcol['busservicid_count'] = busServiceIdlist

    finaldfdcol = finaldfdcol.replace({"busservicid_count": tempdfdict})
    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldfdcol.to_excel(writer_object, startcol=2, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # print(finaldf)
    # E & F column routeid

    routeidlist = dfrefoperator['routeid'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf1 = driver.dataframe1[(driver.dataframe1.type.isin(['routeId']))]
    tempdf1 = tempdf1.drop(['type'], axis=1)
    tempdf1dict = tempdf1.set_index('id')['count'].to_dict()

    # print(tempdf1dict)

    routeidlist = list(map(str, routeidlist))  # casting list to str

    finaldfedcol = pd.DataFrame(routeidlist, columns=['routeId '])
    # print(finaldfdcol)
    finaldfedcol['routeId_count'] = routeidlist
    # print(finaldfedcol)

    finaldfedcol = finaldfedcol.replace({"routeId_count": tempdf1dict})

    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldfedcol.to_excel(writer_object, startcol=4, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # G and H column

    travellist = dfrefoperator['travels'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf2 = driver.dataframe1[(driver.dataframe1.type.isin(['travels']))]
    tempdf2 = tempdf2.drop(['type'], axis=1)
    tempdf2dict = tempdf2.set_index('id')['count'].to_dict()

    # print(tempdfdict)

    finaldf2dcol = pd.DataFrame(travellist, columns=['travels '])
    finaldf2dcol['travels_count'] = travellist

    finaldf2dcol = finaldf2dcol.replace({"travels_count": tempdf2dict})
    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldf2dcol.to_excel(writer_object, startcol=6, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # I and J column

    busTypeIdlist = dfrefoperator['bustypeid'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf3 = driver.dataframe1[(driver.dataframe1.type.isin(['busTypeId']))]
    tempdf3 = tempdf3.drop(['type'], axis=1)
    tempdf3dict = tempdf3.set_index('id')['count'].to_dict()

    # print(tempdfdict)

    finaldf3dcol = pd.DataFrame(busTypeIdlist, columns=['busTypeId '])
    finaldf3dcol['busTypeId_count'] = busTypeIdlist

    finaldf3dcol['busTypeId_count'].replace(tempdf3dict)

    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldf3dcol.to_excel(writer_object, startcol=8, startrow=2, sheet_name='Sheet1', index=False, header=None)

    workbook_object = writer_object.book
    worksheet_object = writer_object.sheets['Sheet1']
    worksheet_object.set_column('A:B', 30)
    worksheet_object.set_column('C:D', 30)
    worksheet_object.set_column('E:F', 30)
    worksheet_object.set_column('G:H', 30)
    worksheet_object.set_column('I:J', 30)

    worksheet_object.write('A2', "Operator ")

    worksheet_object.write('B2', "Unique_count_Operator ")
    worksheet_object.write('C2', "busserviceid ")
    worksheet_object.write('D2', "count_busserviceid ")
    worksheet_object.write('E2', " routeId")
    worksheet_object.write('F2', "count_routeId ")
    worksheet_object.write('G2', " travels")
    worksheet_object.write('H2', "count_travels ")
    worksheet_object.write('I2', " busTypeId")
    worksheet_object.write('J2', "count_busTypeId ")

    worksheet_object.write('A1', "Total_rows_all _col:" + str(dfrefoperator.shape[0]))

    writer_object.save()


def refbustypeIDcount():
    conn = mysql.connector.connect(user=driver.user, password=driver.password, host=driver.host,
                                   database=driver.databasename)
    mycursor = MySQLCursor(conn)
    mycursor.execute(
        'SELECT  DISTINCT(busTypeId),busServiceId,routeId,travels,operator FROM available_trips ORDER BY busTypeId ASC')
    refbustypeID = mycursor.fetchall()
    # dfrefbusserviceid
    # A column
    dfbustypeID = pd.DataFrame(refbustypeID)  # main dataframe
    dfbustypeID.columns = ['bustypeid', 'busservicid', 'routeid', 'travels', 'operator', ]

    dfbustypeID = dfbustypeID.drop_duplicates(subset='bustypeid', keep='first', inplace=False)

    dfbustypeID.replace('', 'BLANK', inplace=True)

    writer_object = pd.ExcelWriter(driver.desktop + '5_refbustypeID.xlsx')

    dfbustypeID['bustypeid'].to_excel(writer_object, startcol=0, startrow=1, sheet_name='Sheet1', index=False)

    # B column
    tempdf = driver.dataframe1[(driver.dataframe1.type.isin(['busTypeId']))]
    tempdf = tempdf.drop(['type'], axis=1)

    tempdfdict = tempdf.set_index('id')['count'].to_dict()
    # print(tempdfdict)
    finaldf = pd.DataFrame()
    listbcol = dfbustypeID['bustypeid'].tolist()

    finaldf = pd.DataFrame({
        'bustypeid': listbcol,
        'Unique_count_bustypeid': listbcol})

    finaldf['Unique_count_bustypeid'].replace(tempdfdict)

    finaldf.to_excel(writer_object, startcol=0, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # C and D column

    busServiceIdlist = dfbustypeID['busservicid'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf = driver.dataframe1[(driver.dataframe1.type.isin(['Busserviceid']))]
    tempdf = tempdf.drop(['type'], axis=1)
    tempdfdict = tempdf.set_index('id')['count'].to_dict()

    # print(tempdfdict)

    finaldfdcol = pd.DataFrame(busServiceIdlist, columns=['busservicid '])
    finaldfdcol['busservicid_count'] = busServiceIdlist

    finaldfdcol = finaldfdcol.replace({"busservicid_count": tempdfdict})
    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldfdcol.to_excel(writer_object, startcol=2, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # print(finaldf)
    # E & F column routeid

    routeidlist = dfbustypeID['routeid'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf1 = driver.dataframe1[(driver.dataframe1.type.isin(['routeId']))]
    tempdf1 = tempdf1.drop(['type'], axis=1)
    tempdf1dict = tempdf1.set_index('id')['count'].to_dict()

    # print(tempdf1dict)

    routeidlist = list(map(str, routeidlist))  # casting list to str

    finaldfedcol = pd.DataFrame(routeidlist, columns=['routeId '])
    # print(finaldfdcol)
    finaldfedcol['routeId_count'] = routeidlist
    # print(finaldfedcol)

    finaldfedcol = finaldfedcol.replace({"routeId_count": tempdf1dict})

    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldfedcol.to_excel(writer_object, startcol=4, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # G and H column

    travellist = dfbustypeID['travels'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf2 = driver.dataframe1[(driver.dataframe1.type.isin(['travels']))]
    tempdf2 = tempdf2.drop(['type'], axis=1)
    tempdf2dict = tempdf2.set_index('id')['count'].to_dict()

    # print(tempdfdict)

    finaldf2dcol = pd.DataFrame(travellist, columns=['travels '])
    finaldf2dcol['travels_count'] = travellist

    finaldf2dcol = finaldf2dcol.replace({"travels_count": tempdf2dict})
    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldf2dcol.to_excel(writer_object, startcol=6, startrow=2, sheet_name='Sheet1', index=False, header=None)

    # I and J column

    operatorlist = dfbustypeID['operator'].tolist()
    # print(travelslist.count('Blank'))                    #LIST1
    # print(len(travelslist))

    tempdf3 = driver.dataframe1[(driver.dataframe1.type.isin(['Operator']))]
    tempdf3 = tempdf3.drop(['type'], axis=1)
    tempdf3dict = tempdf3.set_index('id')['count'].to_dict()

    # print(tempdfdict)

    finaldf3dcol = pd.DataFrame(operatorlist, columns=['operator '])
    finaldf3dcol['operator_count'] = operatorlist

    finaldf3dcol['operator_count'].replace(tempdf3dict)

    # finaldfdcol=pd.DataFrame(travelslist, columns=['travels_count '])

    # print(finaldfdcol)

    finaldf3dcol.to_excel(writer_object, startcol=8, startrow=2, sheet_name='Sheet1', index=False, header=None)

    workbook_object = writer_object.book
    worksheet_object = writer_object.sheets['Sheet1']
    worksheet_object.set_column('A:B', 30)
    worksheet_object.set_column('C:D', 30)
    worksheet_object.set_column('E:F', 30)
    worksheet_object.set_column('G:H', 30)
    worksheet_object.set_column('I:J', 30)

    worksheet_object.write('A2', "bustypeID ")

    worksheet_object.write('B2', "Unique_bustypeID ")
    worksheet_object.write('C2', "busserviceid ")
    worksheet_object.write('D2', "count_busserviceid ")
    worksheet_object.write('E2', " routeId")
    worksheet_object.write('F2', "count_routeId ")
    worksheet_object.write('G2', " travels")
    worksheet_object.write('H2', "count_travels ")
    worksheet_object.write('I2', " operator")
    worksheet_object.write('J2', "count_operator ")

    worksheet_object.write('A1', "Total_rows_all _col:" + str(dfbustypeID.shape[0]))

    writer_object.save()
