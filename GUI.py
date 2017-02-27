#!/user/bin/env python
# encoding: utf-8
#from Tkinter import *
import Tkinter
import MainControl
import FileDialog
import tkFileDialog
import time
import threading
# def resize(ev=None):
#     label.config(font='Helvetica -%d bold' % scale.get())

FILEDIR = 'E:\projects\jinhua'
SOURCEFILE = '\GSM话务量&流量指标(1月).xlsx'
class GUI(object):


    def __init__(self):
        self.mc = MainControl.MainControl()
        self.filename = FILEDIR + SOURCEFILE

    def selectSourceFile(self):
        self.filename = tkFileDialog.askopenfilename(initialdir=FILEDIR)
        self.sourceexcelname.set(self.filename)
        print self.filename

    def selectTargetFile(self):
        self.targetfilename = tkFileDialog.askopenfilename(initialdir=FILEDIR)
        self.targetexcelname.set(self.targetfilename)
        print "targetfile", self.targetfilename

    def selectOutputFile(self):
        self.outputfilename = tkFileDialog.asksaveasfilename(initialdir=FILEDIR)
        self.outputexcelname.set(self.outputfilename)
        print "outputfile", self.outputfilename

    def title(self):
        self.filestate.set('文件载入中')

    def file(self):
        self.mc.loadFile(self.filename, self.targetfilename)
        self.filestate.set('文件已载入')
        self.load['state'] = Tkinter.NORMAL

    def loadFile(self):
        self.filestate.set('文件载入中')
        self.load['state'] = Tkinter.DISABLED
        t = threading.Thread(target=self.file)
        #t.setDaemon(True)
        t.start()


    def old_loadFile(self):
        self.filestate.set('文件载入中')
        #self.filestate.set('文件载入中')
        self.mc.loadFile(self.filename, self.targetfilename)
        self.filestate.set('文件已载入')

    def GUIoutputFile(self):
        b = self.begindate.get()
        e = self.enddate.get()
        print b, e
        self.mc.outputFile(b, e, self.outputfilename)

    def mainLoop(self):
        top = Tkinter.Tk()
        top.title("平均业务量提取工具")
        top.geometry('240x200')

        #引入框架
        frame = Tkinter.Frame(top)
        frame.pack(expand=1, fill=Tkinter.BOTH)

        #选择源文件
        self.sourceexcelname = Tkinter.StringVar()
        sourcefilename = Tkinter.Entry(frame, textvariable=self.sourceexcelname, state='readonly')
        self.sourceexcelname.set("")
        #sourcefilename.pack(expand=1)
        sourcefilename.grid(row=0, column=0, columnspan=2)

        select = Tkinter.Button(frame, text='选择源文件', command=self.selectSourceFile)
        #select.pack()
        select.grid(row=0, column=2, sticky='nsew')

        #选择目标文件
        self.targetexcelname = Tkinter.StringVar()
        targetfilename = Tkinter.Entry(frame, textvariable=self.targetexcelname, state='readonly')
        self.targetexcelname.set("")
        #targetfilename.pack(expand=1, fill=Tkinter.X)
        targetfilename.grid(row=1, column=0, columnspan=2)

        select = Tkinter.Button(frame, text='选择小区名称', command=self.selectTargetFile)
        #select.pack()
        select.grid(row=1, column=2, sticky='nsew')

        # 载入文件
        self.filestate = Tkinter.StringVar()
        filestate = Tkinter.Label(frame, textvariable=self.filestate, fg='SlateGray')
        self.filestate.set('文件未载入')
        filestate.grid(row=2, column=0)
        self.load = Tkinter.Button(frame, text='载入文件', command=self.loadFile)

        self.load.grid(row=2, column=1,columnspan=3)

        #起始日期
        beginlab = Tkinter.Label(frame, text="开始日期")
        beginlab.grid(row=3, column=0)
        self.begindate = Tkinter.StringVar()
        begindateentry = Tkinter.Entry(frame, textvariable=self.begindate)
        begindateentry.grid(row=3, column=1, columnspan=2)

        #结束日期
        endlab = Tkinter.Label(frame, text="结束日期")
        endlab.grid(row=4, column=0)
        self.enddate = Tkinter.StringVar()
        enddateentry = Tkinter.Entry(frame, textvariable=self.enddate)
        enddateentry.grid(row=4, column=1, columnspan=2)

        # 选择输出文件
        self.outputexcelname = Tkinter.StringVar()
        outputfilename = Tkinter.Entry(frame, textvariable=self.outputexcelname)
        self.outputexcelname.set("")
        # targetfilename.pack(expand=1, fill=Tkinter.X)
        outputfilename.grid(row=5, column=0, columnspan=2)

        select = Tkinter.Button(frame, text='选择输出路径', command=self.selectOutputFile)
        # select.pack()
        select.grid(row=5, column=2, sticky='nsew')

        #输出结果
        output = Tkinter.Button(frame, text='输出结果', command=self.GUIoutputFile)
        output.grid(row=6, columnspan=3)
        # quit = Tkinter.Button(frame, text='退出', command=top.quit, activeforeground='white', activebackground='red')
        # quit.grid(row=5, column=1)

        Tkinter.mainloop()


if __name__ == "__main__":
    gui = GUI()
    gui.mainLoop()









