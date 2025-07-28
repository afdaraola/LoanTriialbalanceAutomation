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

#insert single row
    #sql = "insert into customer (name,address) values (%s,%s)"
    #val= ("Ade","78 trembles")
    #mydb.execute(sql,values)
    #conn.commit()

#insert multiple row 
sql = "insert into customer (name,address) values (%s, %s)"
values= [
        ('Ade', '78 trembles'),
        ('john', '48 laval'),
        ('figo', '56 montreal'),
        ('rice', '78 jean coutu'),
        ('palmer', '78 london')
        ]

mydb.executemany(sql,values)

conn.commit()

print(str(mydb.rowcount) +" record count")