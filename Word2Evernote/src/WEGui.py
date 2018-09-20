# -*- coding: utf-8 -*-
from Tkinter import *
import tkFileDialog
from WordExtract import *
from EvernoteController import *
from Oauth import *

class WEGui(object):
    def __init__(self, initdir=None):
        '''构造函数，说明版本信息'''
        self.top = Tk()
        self.top.title("印象笔记推送工具")

        self.inputdir = StringVar(self.top)
        self.inputfm = Frame(self.top, width=80)

        self.space1 = Label(self.top, text='', height=1)
        self.space1.pack()

        self.inputstr = Label(self.inputfm, font=('宋体', 10, 'bold'), text='选择文件',width=10,height=2)
        self.inputentry = Entry(self.inputfm, borderwidth=2, width = 60,bg='#F8F8FF', textvariable=self.inputdir)
        self.inputentry.config(state=DISABLED)
        self.inputbrower = Button(self.inputfm, text='浏览',bg='#ADD8E6', width = 6,command=self.getInputFileName,font=('宋体', 10, 'bold'), activeforeground = 'white', activebackground = '#00BFFF')
        self.inputstr.pack(side=LEFT)
        self.inputentry.pack(side=LEFT)
        self.inputbrower.pack(side=LEFT)
        self.inputfm.pack_forget()

        self.tnname = StringVar(self.top)
        self.tnotenamefm = Frame(self.top, width=80)
        self.tnotename = Label(self.tnotenamefm, font=('宋体', 10, 'bold'), text='笔记本名称', width=12, height=2)
        self.tnotenameentry = Entry(self.tnotenamefm, borderwidth=2, width=60, bg='#F8F8FF', textvariable=self.tnotename)
        self.tnotename.pack(side=LEFT)
        self.tnotenameentry.pack(side=LEFT)
        self.tnotenamefm.pack_forget()

        self.login = StringVar(self.top)
        self.longinfm = Frame(self.top, width=80)
        self.loginname = Label(self.longinfm, font=('宋体', 10, 'bold'), text='印象笔记用户名', width=12, height=2)
        self.loginnameentry = Entry(self.longinfm, borderwidth=2, width=25, bg='#F8F8FF', textvariable=self.loginname)
        self.loginpasswd = Label(self.longinfm, font=('宋体', 10, 'bold'), text='密码', width=5, height=2)
        self.loginpasswdentry = Entry(self.longinfm, borderwidth=2, width=25, bg='#F8F8FF', textvariable=self.loginpasswd, show='*')
        self.loginButton = Button(self.longinfm, text='登录', bg='#ADD8E6', width=6, command=self.loginEvent,font=('宋体', 10, 'bold'), activeforeground='white', activebackground='#00BFFF')
        self.loginname.pack(side=LEFT)
        self.loginnameentry.pack(side=LEFT)
        self.loginpasswd.pack(side=LEFT)
        self.loginpasswdentry.pack(side=LEFT)
        self.loginButton.pack(side=LEFT)
        self.longinfm.pack()

        self.extract = Button(self.top,text='信息推送', font=('宋体', 11, 'bold'),bg='#ADD8E6', command=self.extractInfo, activeforeground='white',activebackground='#00BFFF',width=60)
        self.extract.pack_forget()

        self.space2 = Label(self.top, text='', height=1)
        self.space2.pack()

        self.dirfm = Frame(self.top)
        self.dialog = Label(self.dirfm, text="   处理日志",justify=LEFT)
        self.dirsb = Scrollbar(self.dirfm)
        self.dirsb.pack(side=RIGHT, fill=Y)
        self.dirs = Listbox(self.dirfm, height=15, width=80, yscrollcommand=self.dirsb.set)
        self.dirsb.config(command=self.dirs.yview)
        self.dialog.pack()
        self.dirs.pack(side=LEFT, fill=X, ipadx=20)
        self.dirfm.pack()

        self.token = None

    def getInputFileName(self):
        dirname = tkFileDialog.askopenfilenames(filetypes=[('Word 97-2003 文档', '.doc'), ('Word 文档', '.docx')])[0]
        self.inputdir.set(dirname)

    def loginEvent(self):

        self.dirs.insert(END, '登录中， 用户名：' + self.loginnameentry.get() + ' ......')
        self.dirs.see(self.dirs.index(END))
        self.dirs.update()

        key = 'key'
        secret = 'secret'
        account = self.loginnameentry.get()
        passwd = self.loginpasswdentry.get()
        self.token = Oauth(key, secret, account, passwd).oauth()
        if not self.token:
            self.dirs.insert(END, '登录失败，请确保用户名密码正确、网络正常、完成授权')
            self.dirs.see(self.dirs.index(END))
            self.dirs.update()
        else:
            self.dirs.insert(END, '登录成功， 用户名：' + self.loginnameentry.get())
            self.dirs.see(self.dirs.index(END))
            self.dirs.update()
            self.top.title("印象笔记推送工具--" + str(self.loginnameentry.get()))
            self.longinfm.pack_forget()
            self.dirfm.pack_forget()
            self.space1.pack_forget()
            self.space2.pack_forget()
            self.inputfm.pack()
            self.tnotenamefm.pack()
            self.extract.pack()
            self.space2.pack()
            self.dirfm.pack()
        return

    def extractInfo(self):
        if not self.inputdir.get():
            self.dirs.insert(END, '请选择要抽取信息的word文档')
            self.dirs.see(self.dirs.index(END))
            self.dirs.update()
            return

        self.dirs.insert(END, '=========================开始=========================')
        self.dirs.see(self.dirs.index(END))
        self.dirs.update()

        self.inputbrower.config(state=DISABLED)
        self.extract.config(stat=DISABLED)

        temp = WordExtract(self.inputdir.get()).getDocument()
        if not temp or not temp[1]:
            self.dirs.insert(END, 'word文档抽取失败，请检查当前文档是否为待抽取文档')
            self.dirs.see(self.dirs.index(END))
            self.dirs.update()

        else:
            self.dirs.insert(END, 'word文档抽取完成')
            self.dirs.see(self.dirs.index(END))
            self.dirs.update()
            nb, keyVal = temp

            ec = EvernoteController(self.token)
            if self.tnotenameentry.get():
                nb = str(self.tnotenameentry.get()).strip()
            guid = ec.createNotebook(nb)

            if not guid:
                self.dirs.insert(END, '连接印象笔记记事本失败，请重试')
                self.dirs.see(self.dirs.index(END))
                self.dirs.update()
                self.inputbrower.config(state=NORMAL)
                self.extract.config(stat=NORMAL)
                return

            for (k, v) in keyVal.items():
                if ec.create_note(k,  v, guid):
                    self.dirs.insert(END, '已推送 ' + str(k.encode('utf-8')) + ' 到记事本: ' + str(nb.encode('utf-8')))
                    self.dirs.see(self.dirs.index(END))
                    self.dirs.update()
                else:
                    self.dirs.insert(END, '推送' + str(k.encode('utf-8')) + '到记事本: ' + str(nb.encode('utf-8')) + ' 失败')
                    self.dirs.see(self.dirs.index(END))
                    self.dirs.update()

            self.dirs.insert(END, '文档  ' + str(nb.encode('utf-8')) + ' 已并推送到 ' + str(self.loginnameentry.get()) + ' 的印象笔记')
            self.dirs.see(self.dirs.index(END))
            self.dirs.update()
            self.dirs.insert(END, '=========================结束========================')
            self.dirs.see(self.dirs.index(END))
            self.dirs.update()

        self.inputbrower.config(state=NORMAL)
        self.extract.config(stat=NORMAL)
