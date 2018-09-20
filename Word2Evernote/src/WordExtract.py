# -*- coding: utf-8 -*-
import os
from win32com.client import Dispatch

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
