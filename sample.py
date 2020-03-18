import mysql.connector
from mysql.connector.cursor import MySQLCursor


def display():
    conn = mysql.connector.connect(user='viplav', password='password', host='127.0.0.1' ,database='stocklaundry')
    mycursor = MySQLCursor(conn)

    mycursor.execute('SELECT  `Item_name`, `Quantity`FROM ` balance_stock` WHERE Date= "2019-06-19"')
    sbothval = mycursor.fetchall()

    sbothvall = dict(sbothval)

    print(sbothvall)

display()





