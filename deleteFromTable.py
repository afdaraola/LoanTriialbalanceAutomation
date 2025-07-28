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


#This is to prevent SQL injections, which is a common web hacking technique to destroy or misuse your database.
#The mysql.connector module uses the placeholder %s to escape values in the delete statement:

myquery = "delete from customer where id = %s"
val = 2

mydb.execute(myquery,val)

conn.commit()

print(str(mydb.rowcount) + " Record delete successfuly")