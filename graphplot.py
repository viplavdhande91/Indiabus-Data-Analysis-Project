from openpyxl import Workbook
import mysql.connector
from openpyxl.styles import Font
from itertools import zip_longest
from mysql.connector.cursor import MySQLCursor
from openpyxl.styles import PatternFill
import driver
import pandas as pd
from collections import Counter
import numpy as np
import matplotlib.pyplot as plt; plt.rcdefaults()
import matplotlib.mlab as mlab
import matplotlib.gridspec as gridspec
import operator
import matplotlib.pyplot as plt


def Graphdraw(val):
  # print(val)
   val =list(val.values())
   val=val[1:]        #slccing to remove blank entry

   newval=dict(Counter(val))  #counting ocuurence of counts
   #print(newval)
   list1=newval.values()

   newnewval = dict(Counter(list1))

   newnewval = dict(sorted(newnewval.items(), key=operator.itemgetter(0))) #sort dict by keys
   #print(newnewval)

   xx = list(newnewval.keys())
   y_pos = np.arange(len(xx))
   performance = list(newnewval.values())

   bars=plt.bar(y_pos, performance,width=0.5,align='center',alpha=1)
   plt.xticks(y_pos, xx)
   plt.ylabel('Counts')
   plt.title('BusService ID Counts Ocurrence')

   for bar in bars:
       yval = bar.get_height()
       plt.text(bar.get_x(), yval + .005, yval)




   plt.savefig(driver.desktop+'graphs\\busserviceidcount.svg',format='svg', dpi=1200,bbox_inches='tight')
   #plt.show()

#test()
 #  num1.remove("")
 #  num2.pop(0)









#  plt.savefig('C:\\Users\\AdminPC\\Desktop\\OUTPUTFOLDER\\Associated_counts\\busservicegraph.png')


