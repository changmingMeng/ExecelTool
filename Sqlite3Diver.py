# encoding: utf-8
import sqlite3


# import time

# createtabsql1 = "create table if not exists scriptdata(id integer primary key autoincrement, name varchar(128), info varchar(128))"
class Sqlite3Driver(object):
    '''''
    The Sqlite3Driver class use to write the script data that parse from Excel file
    '''

    def __init__(self, dbfile, tabledesc):
        self.tablename = tabledesc[0]
        self.tablefield = tabledesc[1]
        self.dbfile = dbfile

    def cerateDB(self):
        createlist = ["create table if not exists ", self.tablename, "(id integer primary key autoincrement, ",
                      self.tablefield, ")"]
        createsql = "".join(createlist)
        self.conn = sqlite3.connect(self.dbfile)
        self.conn.isolation_level = None
        # self.conn.execute("drop table if exists " + self.tablename)  ###delete the eixst table
        self.conn.execute(createsql)  ####create new table
        # conn.execute("delete from " + tablename) ####delete all the recoreds
        return

    def execDB(self, execsql):
        #print execsql
        self.conn.execute(execsql)
        self.conn.commit()
        return

    def getResult(self, selectsql):
        self.cur = self.conn.cursor()
        self.cur.execute(selectsql)
        self.res = self.cur.fetchall()
        self.cur.close()
        return self.res

    def getCount(self):
        return len(self.res)

    def closeDB(self):
        self.cur.close()
        self.conn.close()


'''''
The example for using the DBDriver class
'''
if __name__ == '__main__':
    # dbfile = ":memory:"
    # tabledesc = ("scriptdata", "name varchar(128), info varchar(128)")
    # insertsql = "insert into scriptdata(name,info) values ('zhaowei1', 'only a test')"
    # selectsql = "select * from scriptdata"

    dbfile = "E:\\projects\\jinhua\\datebase.txt"
    tabledesc = ("exceldata", "name varchar(128), date date, erl float, data float")
    selectsql = "select * from exceldata"

    dbd = Sqlite3Driver(dbfile, tabledesc)
    dbd.cerateDB()
    #dbd.execDB(insertsql)
    res = dbd.getResult(selectsql)
    rows = dbd.getCount()
    dbd.closeDB()

    print 'row:', rows
    for line in res:
        print line
        # for col in line:
        #     print col