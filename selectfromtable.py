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

#myquery = "select * from customer"

myquery = "select * from customer where name = %s order by 1 desc"
nam = "ade"

mydb.execute(myquery,nam)

#fetchone only 
    #result = mydb.fetchone()
    #print(result)

#fetchall 
results = mydb.fetchall()
 
for x in results:
    print(x)