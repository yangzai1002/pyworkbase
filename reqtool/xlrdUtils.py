import xlrd
class xlrdUtils():
    def __init__(self,path):
        self.path=path
        self.book=xlrd.open_workbook(self.path)
        self.table=self.book.sheet_by_index(0)
    def printSheet1(self):
        self.usedRow=self.table.nrows
        self.tablename=self.table.name
        self.tablecol=self.table.ncols
        print(self.tablename+" row:"+str(self.usedRow)+"cloumn:"+str(self.tablecol))

if __name__=="__main__":
    path="F:/pyworkbase/reqtool/Requirement_specification.xls"
    myexcel=xlrdUtils(path)
    myexcel.printSheet1()
