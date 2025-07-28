import pymysql

try:
    conn = pymysql.connect(
        host ="localhost",
        user="root",
        password="W3lc0m3@123",
        database="developer"
    )
except Exception as e:
    print(e)

if conn:
    print("Connected sucessfully")

mydb = conn.cursor()

mycursor = mydb.execute("create table customer(id int auto_increment primary key, name varchar(255), address varchar(255))")

