import mysql.connector
from pymongo import MongoClient

MYSQL_AUTH_PLUGIN = "mysql_native_password"

def helloWorld():
    print("Hello World!")

# 获取MySQL连接
def getMySQLConn(user,password,host,port,database):
    conn = mysql.connector.connect(user=user, password=password,
        host=host, port=port,
        database=database, 
	    auth_plugin = MYSQL_AUTH_PLUGIN,
        use_unicode=True)
    return conn

def getMongoDBClient(ip,username,password,authSource,authMechanism):
    client = MongoClient(ip,
                     username=username,
                     password=password,
                     authSource=authSource,
                     authMechanism=authMechanism)
    return client