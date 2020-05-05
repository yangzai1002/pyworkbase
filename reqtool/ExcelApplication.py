from win32com.client import Dispatch
class myExcel():
    def __init__(self):
        self.myexcel=Dispatch('Excel.Application')
        self.myexcel.Visible = 0
        self.myexcel.DisplayAlerts = 0
    def AddBook(self):
        self.wb = self.myexcel.Workbooks.Add()
        self.wt = self.wb.Worksheets("sheet1")
        return self.wb,self.wt
    def OpenBook(self,path,sheetname):
        self.wb=self.myexcel.Workbooks.Open(path)
        self.wt=self.wb.Worksheets(sheetname)
        return self.wb,self.wt
    def getSheets(self,path):
        self.wb=self.myexcel.Workbooks.Open(path)
        self.wts=self.wb.Worksheets
        return self.wb,self.wts
    def unionFormat(self,ws,range):
        ws.Range(range).RowHeight = 24
        ws.Range(range).ColumnWidth = 30
        ws.Range(range).Font.Name = "微软雅黑"
        ws.Range(range).Font.Size = 10
    def makeColor(self,colornumber):
        self.myexcel.Selection.Interior.Color = colornumber
    def Quit(self):
        self.myexcel.Quit()
    def selection(self):
        return self.myexcel.Selection