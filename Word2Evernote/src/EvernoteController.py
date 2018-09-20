# -*- coding: utf-8 -*-
from evernote.api.client import EvernoteClient
import evernote.edam.type.ttypes as Types

class EvernoteController(object):
    def __init__(self, token):
        self.token = token
        self.client = EvernoteClient(token=self.token)
        self.userStore = self.client.get_user_store()
        self.noteStore = self.client.get_note_store()
        self.username = self.userStore.getUser().username

    ##　创建记事本，如果已存在，不创建, 返回guid
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



