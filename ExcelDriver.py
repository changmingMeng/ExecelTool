# encoding: utf-8
import xlrd
import xlwt
import datetime
import Sqlite3Diver
import utils
import re

class ExcelDriver(object):


    def __init__(self):
        #self.fileName = u"E:\projects\jinhua\GSM话务量&流量指标(1月).xlsx"
        self.filetype = ""
        pass

    def old_getTargets(self, fileName):
        book = xlrd.open_workbook(fileName)
        self.targets = []
        for r in range(book.sheet_by_index(0).nrows):
            self.targets.append(book.sheet_by_index(0).row(r)[0].value)
        print self.targets
        return self.targets

    #相比于old_getTargets，使用了with...as..语句来代替try...except语句
    def getTargets(self, fileName):
        self.targets = []
        with xlrd.open_workbook(fileName) as book:
            for r in range(book.sheet_by_index(0).nrows):
                self.targets.append(book.sheet_by_index(0).row(r)[0].value)
            print self.targets
            return self.targets

    def readFile(self, fileName):
        if re.search('GSM', fileName):
            self.filetype = '2G'
        elif re.search('WCDMA', fileName):
            self.filetype = '3G'
        else:
            self.filetype = '4G'
        print fileName, self.filetype
        self.book = xlrd.open_workbook(fileName)

    def writeFile(self, book, rowNames, searchdates, fileName = "E:\\projects\\jinhua\\results.xls"):
        file = xlwt.Workbook(encoding='utf-8')
        table = file.add_sheet('sheet1', cell_overwrite_ok=True)
        #表格第一行输出日期
        j = 0
        for date in searchdates:
            table.write(0, j, date)
            j += 1
        #表格第二行输出标题
        j = 0
        for name in rowNames:
            table.write(1,j, name)
            j += 1
        #表格第三行开始输出数据
        i = 2
        for row in book:
            j = 0
            print "row =", row
            for v in row[0]:
                print "v= ", v
                # if type(v) is not str:
                #     table.write(i, j, str(v))
                table.write(i,j,v)
                j += 1
            i += 1
        print fileName
        file.save(fileName)

    def pickResult(self):
        print self.filetype
        if self.filetype == '2G':
            print 'pickResult2G'
            return self.pickResult2G()
        elif self.filetype == '3G':
            print 'pickResult3G'
            return self.pickResult3G()
        else:
            print 'pickResult4G'
            return self.pickResult4G()

    def pickResult2G(self):
        # 遍历第二个表单
        sheet = self.book.sheet_by_index(1)
        self.results = []
        ut = utils.Utils()
        for r in range(sheet.nrows):
            if sheet.cell_value(r,1) in self.targets:
                name = sheet.cell_value(r,1)
                dateTemp = sheet.cell_value(r,0)
                old_date = xlrd.xldate_as_tuple(dateTemp,0)[0:3]
                date = ut.tuple2Sqlite3Timestring(old_date)
                erl = sheet.cell_value(r,5)
                data = sheet.cell_value(r,6)
                # self.Save2DB(name, date, erl, data)
                print name, date, erl, data
                self.results.append((name, date, erl, data))
        return self.results

    def pickResult3G(self):

        self.results = []
        ut = utils.Utils()
        for i in (0, 2):
            sheet = self.book.sheet_by_index(i)
            for r in range(sheet.nrows):
                if sheet.cell_value(r,3) in self.targets:
                    name = sheet.cell_value(r,3)
                    dateTemp = sheet.cell_value(r, 0)
                    old_date = xlrd.xldate_as_tuple(dateTemp, 0)[0:3]
                    date = ut.tuple2Sqlite3Timestring(old_date)
                    erl = sheet.cell_value(r, 6)
                    data = sheet.cell_value(r, 7)+sheet.cell_value(r, 8)+sheet.cell_value(r, 9)+sheet.cell_value(r, 10)
                    data /= (1024*1024)
                    print name, date, erl, data
                    self.results.append((name, date, erl, data))
        return self.results

    def pickResult4G(self):
        sheet = self.book.sheet_by_index(1)
        self.results = []
        ut = utils.Utils()
        for r in range(sheet.nrows):
            if sheet.cell_value(r, 4) in self.targets:
                name = sheet.cell_value(r, 4)
                dateTemp = sheet.cell_value(r, 0)
                old_date = xlrd.xldate_as_tuple(dateTemp, 0)[0:3]
                date = ut.tuple2Sqlite3Timestring(old_date)
                erl = 0
                data = sheet.cell_value(r, 5)
                print "4g",name, date, erl, data
                self.results.append((name, date, erl, data))
        return self.results


if __name__ == "__main__":
    startTime = datetime.datetime.now()
    print "start time:", startTime

    ed = ExcelDriver()
    ed.readFile(u"E:\projects\jinhua\WCDMA话务量&流量指标(201701月).xlsx")
    ed.getTargets("E:\\projects\\jinhua\\targets3G.xlsx")
    ed.pickResult3G()

    # book = xlrd.open_workbook(u"E:\projects\jinhua\GSM话务量&流量指标(1月).xlsx")
    # print "表单数量:", book.nsheets
    # print "表单名称:", book.sheet_names()
    # # 获取第1个表单
    # sh = book.sheet_by_index(0)
    # print u"表单 %s 共 %d 行 %d 列" % (sh.name, sh.nrows, sh.ncols)
    # print "第二行第三列:", sh.cell_value(1, 1)
    # print "".join(sh.row(0))
    # print "".join(sh.row(1))
    endTime = datetime.datetime.now()
    print "end time:", endTime
    print "run time:", (endTime - startTime)
    # 遍历所有表单
    # for s in book.sheets():
    #     for r in range(s.nrows):
    #         # 输出指定行
    #         print s.row(r)