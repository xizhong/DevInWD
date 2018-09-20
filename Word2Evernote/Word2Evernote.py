# -*- coding: utf-8 -*-
import requests
import urllib
from evernote.api.client import EvernoteClient
import evernote.edam.type.ttypes as Types
import hashlib
import os
from win32com.client import Dispatch
from Tkinter import *
import tkFileDialog
#
# import sys
# defaultencoding = 'utf-8'
# if sys.getdefaultencoding() != defaultencoding:
#     reload(sys)
#     sys.setdefaultencoding(defaultencoding)

class Oauth(object):
    def __init__(self, consumerKey, consumerSecret, account, passwd, sandbox=True, isInternational = False):
        self.consumerKey = consumerKey
        self.consumerSecret = consumerSecret
        self.accout = account
        self.passwd = passwd
        if sandbox:
            self.host = 'sandbox.evernote.com'
        elif isInternational:
            self.host = 'app.evernote.com'
        else:
            self.host = 'app.yinxiang.com'

        # self.host = 'app.yinxiang.com'

    def __get_tmp_token(self):
        params = {
            'oauth_callback': '127.0.0.1',
            'oauth_consumer_key': self.consumerKey,
            'oauth_signature': self.consumerSecret,
            'oauth_signature_method': 'PLAINTEXT',
        }

        r = requests.get('https://%s/oauth' % self.host, params=params)
        if not 'oauth_token' in r.text:
            self.tmpOauthToken = None
        else:
            self.tmpOauthToken = str(dict(item.split('=', 1) for item in urllib.unquote(r.text).split('&'))['oauth_token'])

    def oauth(self):
        if os.path.isfile('C:\\tmp_wordextract'):
            with open('C:\\tmp_wordextract', 'r') as f:
                apDict = {}
                for line in f.readlines():
                    if line.replace('\n', ''):
                        key, val = line.split("####", 1)
                        apDict[key] = val
            if self.__get_md5() in apDict.keys():
                return apDict[self.__get_md5()]
        self.__get_tmp_token()
        self.__get_ver()
        token_ = self.__get_token()
        if token_:
            with open('C:\\tmp_wordextract', 'a') as f:
                f.write('\n' + self.__get_md5() + "####" + str(token_[0]))
            return token_[0]
        else:
            return None

    def __get_login_info(self):
        account = self.accout
        password = self.passwd
        return account, password

    def __get_md5(self):
        return  hashlib.md5(str(self.accout) + str(self.passwd)).hexdigest()

    def __get_ver(self):
        if not self.tmpOauthToken:
            self.verifier = None
            return
        account, password = self.__get_login_info()
        access = {
            'authorize': 'Authorize',
            'oauth_token': self.tmpOauthToken,
            'username': account,
            'password': password,
        }
        r = requests.post('https://%s/OAuth.action'%self.host, data = access)
        if 'oauth_verifier' in r.url:
                self.verifier = dict(item.split('=', 1) for item in r.url.split('?')[-1].split('&'))['oauth_verifier']
        else:
            self.verifier = None

    def __get_token(self):
        if not self.verifier:
            return None
        payload = {
            'oauth_consumer_key': self.consumerKey,
            'oauth_token': self.tmpOauthToken,
            'oauth_verifier': self.verifier,
            'oauth_signature': self.consumerSecret,
            'oauth_signature_method': 'PLAINTEXT',
        }
        r = requests.get('https://%s/oauth'%self.host, params = payload)

        if not ('oauth_token' in r.text and 'edam_expires' in r.text):
            raise Exception('Token Not Found')
        return (dict(item.split('=',1) for item in urllib.unquote(r.text).split('&'))['oauth_token'],
            dict(item.split('=',1) for item in urllib.unquote(r.text).split('&'))['edam_expires'], self.host)

class EvernoteController(object):
    def __init__(self, token):
        self.token = token
        self.client = EvernoteClient(token=self.token)
        self.userStore = self.client.get_user_store()
        self.noteStore = self.client.get_note_store()
        self.username = self.userStore.getUser().username

    ##　创建纪事本，如果已存在，不创建, 返回guid
    def createNotebook(self, notebookName):
        notebookDict= self.getNotebook()
        if notebookName in notebookDict.keys():
            guid = notebookDict[notebookName]
        else:
            try:
                notebook = Types.Notebook()
                notebook.name = notebookName.encode('utf-8')
                notebook = self.noteStore.createNotebook(notebook)
                guid = notebook.guid
            except Exception, e:
                guid = None
                print(e.message)
        return guid

    def create_note(self, noteName, content, notebookGuide):
        if not noteName or not content or not noteName:
            return False
        note = Types.Note()
        note.title = noteName.encode('utf-8')
        note.content = '<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE en-note SYSTEM "http://xml.evernote.com/pub/enml2.dtd">'
        note.content += '<en-note>' + content.encode('utf-8') + '</en-note>'
        note.notebookGuid = notebookGuide
        try:
            self.noteStore.createNote(note)
            return True
        except Exception, e:
            return False

    def getNotebook(self):
        # tt = {nb.name: nb.guid for nb in self.noteStore.listNotebooks()}
        return {nb.name.decode('utf-8'): nb.guid for nb in self.noteStore.listNotebooks()}

class WordExtract(object):
    def __init__(self, path):
        self.docPath = path.strip()

    # def getFileList(self):
    #     filelist = list()
    #     if not os.path.isdir(self.docPath) and (self.docPath.endswith('doc') or self.docPath.endswith('docx')):
    #         filelist.append(self.docPath)
    #     else:
    #         if not self.docPath.endswith(os.path.sep):
    #             self.docPath += os.path.sep
    #         for filename in os.listdir(self.docPath):
    #             filepath = self.docPath + filename
    #             if os.path.isfile(filepath) and not filename.startswith('~') and (
    #                     filename.endswith('docx') or filename.endswith('doc')):
    #                 filelist.append(filepath)
    #     if len(filelist) == 0:
    #         return None
    #     else:
    #         return filelist

    def getDocument(self):
        word = Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = ''
        try:
            doc = word.Documents.Open(FileName=self.docPath)
        except Exception as e:
            word.Quit()
        docname = os.path.split(self.docPath)[-1].split('.')[0]
        if not doc:
            return None
        index = 0
        contentKVDict = dict()
        key = None
        val = None
        while index < doc.Paragraphs.Count:
            par = doc.Paragraphs[index]
            content = par.Range.text.replace('\r', '')
            if not content:
                index += 1
                continue
            fontname, fontsize = par.Range.Font.Name, par.Range.Font.Size
            if fontname == '黑体'.decode('utf-8') and int(fontsize) == 14:
                if not key:
                    key = content
                    val = ''
                else:
                    contentKVDict[key] = val
                    key = content
                    val = ''
            if fontname == '宋体'.decode('utf-8') and int(fontsize) == 16:
                if val:
                    val += content
                else:
                    val = content
            index += 1
        if key:
            contentKVDict[key] = val
        word.Quit()
        return docname, contentKVDict

    def delTmpDoc(self):
        pass


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

if __name__ == '__main__':
    WEGui()
    mainloop()






