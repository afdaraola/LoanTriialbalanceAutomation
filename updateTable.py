import pymysql

try:
    conn = pymysql.connect(host = "localhost",
                           user="root",
                           password="W3lc0m3@123",
                           database="developer"
                           )
except Exception as e:
    print(e)


if conn:
    print("Connect is successful")

mydb = conn.cursor()  

#sql = "Update customer set address = '2 animashaun' where id = 3"

sql = "Update customer set address = %s where id = %s"
vars = ("2 london", 3)


mydb.execute(sql,vars)

conn.commit() 

print(mydb.rowcount)