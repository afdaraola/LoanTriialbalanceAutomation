
import pymysql


try: 
     conn = pymysql.connect(
            host="localhost", 
            user="root", 
            password="W3lc0m3@123", 
            database="mysql"
            )  
except Exception as e:
    print("failed to establish connection..")

if conn:
     print ("connection successful") 

mydb = conn.cursor()

#mydb.execute("SHOW DATABASES")
mydb.execute("SHOW TABLES")

for x in mydb:
    print(x)




