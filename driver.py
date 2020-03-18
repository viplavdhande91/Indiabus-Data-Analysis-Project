from openpyxl import Workbook
import mysql.connector
from openpyxl.styles import Font
from itertools import *
from mysql.connector.cursor import MySQLCursor
import errorexcelfile
import parentmain
import module1
import module2
import module3
import time
import graphplot
import os
desktop = os.path.normpath(os.path.expanduser("~/Desktop"))
desktop =desktop + '\\OUTPUTFOLDER\\'
start = time.time()



user = str('viplav')
password = str('password')
host = str('127.0.0.1')
databasename = str('test')

dataframe1=module2.Combinedssociatedcounts2()


if __name__ == "__main__":
    parentmain.Myfunc()
    errorexcelfile.Sample()
    errorexcelfile.Sourceerror()
    errorexcelfile.Destinationerror()
    errorexcelfile.Travelserror()
    dictbusservice = module1.Associatedcounts()
    module1.Associatedcounts2()
    module1.Associatedcounts3()
    module2.dataframewrite(dataframe1)
    graphplot.Graphdraw(dictbusservice)     #we pass value for graph plotting
    module3.refbusserviceidcount()
    module3.reftravelscount()
    module3.refrouteIDcount()
    module3.refoperatorcount()
    module3.refbustypeIDcount()

    end = time.time()
    newval = end - start
    print("Execution time :" + '' + str(newval) + 'seconds')



   # print("Enter database name:")






