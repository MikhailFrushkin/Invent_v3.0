import peewee


user = 'root'
password = 'root'
db_name = 'cells'

dbhandle = peewee.SqliteDatabase('Data/mydatabase.db')