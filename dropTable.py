import pymysql

try:
    conn = pymysql.connect(host = "localhost",
                           user="root",
                           password="W3lc0m3@123",
                           database="mysql"
                           )
except Exception as e:
    print(e)


if conn:
    print("Connect is successful")

mydb = conn.cursor() 

sql = "drop table if exists customers"

mydb.execute(sql)

print("Table drop successfully")

