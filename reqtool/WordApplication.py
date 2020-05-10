from win32com.client import Dispatch
class myWord():
    def __init__(self,path):
        self.ms_word=Dispatch('word.Application')
        self.ms_wordVisible = 0
        self.ms_word.DisplayAlerts = 0
        self.doc=self.ms_word.Documents.Open(path)
    def AcceptRevision(self):
        self.doc.AcceptAllRevisions()
    def ClearFormat(self):
        self.doc.Select()
        self.ms_word.Selection.ClearFormatting()
        self.ms_word.ActiveWindow.View.RevisionsFilter.Markup = 0
        self.ms_word.ActiveWindow.View.RevisionsFilter.View = 0
    def getTables(self):
        return self.doc.Tables
    def getPara(self):
        return self.doc.Paragraphs
    def Save(self):
        self.doc.Save()
    def delTable(self):
        self.Table=self.doc.Tables
        length=len(self.Table)
        for i in range(length):
                self.Table[i].Select()
                self.ms_word.Selection.Cut()
    def Close(self):
        self.doc.Close()
    def Quit(self):
        self.ms_word.Quit()
