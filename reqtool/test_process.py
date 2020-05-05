import os,openpyxl
from win32com.client import Dispatch
class word_process():
    def __init__(self):
        pass
    def catch_word(self):
        self.Path=os.getcwd()
        os.chdir(self.Path)
        self.k=1
        self.file=[0 for self.j in range(40)]
        print(10*'-'+"文档目录"+10*'-')
        try:
            for filename in os.listdir(self.Path):
                print(str(self.k)+':'+filename+'\n')
                self.file[self.k]=filename
                self.k=self.k+1
        except Exception as err:
            print("An exception happend:"+str(err))
            print("文件个数超过40个")
        print("请输入需要操作的文档编号")
        try:
            self.numfile=input()
        except Exception as err:
                print("An exception happend:"+str(err))
        self.name=self.file[int(self.numfile)]
        self.PATH=os.path.join(self.Path,self.name)
        return self.PATH
    def lanch_word(self): # 启动word，不显示修订记录
        self.PATH=self.catch_word()
        self.ms_word=Dispatch('Word.Application')
        self.ms_word.Visible=0
        self.ms_word.DisplayAlerts=0
        self.doc=self.ms_word.Documents.Open(self.PATH)
        self.ms_word.ActiveWindow.View.RevisionsFilter.Markup=0
        self.ms_word.ActiveWindow.View.RevisionsFilter.View=0
        return self.doc
    def lanch_word_1(self): #启动word并打开目标文档进行清除格式和接受修订操作
        self.PATH=self.catch_word()
        self.ms_word=Dispatch('Word.Application')
        self.ms_word.Visible=0
        self.ms_word.DisplayAlerts=0
        self.doc=self.ms_word.Documents.Open(self.PATH)
        self.doc.AcceptAllRevisions()
        self.doc.Select()
        self.ms_word.Selection.ClearFormatting()
        self.ms_word.ActiveWindow.View.RevisionsFilter.Markup=0
        self.ms_word.ActiveWindow.View.RevisionsFilter.View=0
        self.doc.Save()
        return self.doc
    def auto_number_case(self,caselabel): #自动排序函数
        self.doc_1=self.lanch_word()
        self.caselable=caselabel
        self.k=1
        self.number_tab=len(self.doc_1.tables)
        for i in range(self.k, self.number_tab):
            if "Case" in self.doc_1.Tables[i].Rows[0].Cells[0].Range.Text:
                self.text=self.doc_1.Tables[i].Rows[0].Cells[1].Range.Text.splitlines()
                self.doc_1.Tables[i].Rows[0].Cells[1].Select()
                if 1<=self.k<=9:
                    self.text[0]="["+self.caselable+"-000"+str(self.k) +"]"
                elif 10<=self.k<=99:
                    self.text[0]="["+self.caselable+"-00"+str(self.k)+ "]"
                elif 100<=self.k<=1000:
                    self.text[0]="["+self.caselable+"-0"+str(self.k)+ "]"
                self.doc_1.Tables[i].Rows[0].Cells[1].Range.Text='\n'.join(self.text)
                self.k=self.k+1
        self.ms_word.Quit()
    def extract_case(self):
        # self.doc_2=lanch_word()
        self.xlApp = Dispatch('Excel.Application')
        self.xlBook = self.xlApp.Workbooks.Add()
        self_word=self.lanch_word_1()
        self.m=1
        self.row=1
        self.number_tables=len(self_word.tables)
        self.number_case=0 #总用例数
        self.number_step=0 #总的步数
        self.wt=self.xlBook.Worksheets("sheet1")
        try:
            for i in range(self.m,self.number_tables):
                if "Case" in self_word.Tables[i].Rows[0].Cells[0].Range.Text:#判断表格是不是用例表格
                    #self.length1=len(self_word.Tables[i].Rows)
                    '''if "Comment" in self_word.Tables[i].Rows[self.length1-1].Cells[0].Range.Text:# 判断表格末尾是不是有注释的两行
                        self.length=len(self_word.Tables[i].Rows)-2
                    else:
                        self.length=self.length1
                    # 计算总步数
                    self.number_step=self.number_step+self.length'''
                    self.text1=self_word.Tables[i].Rows[0].Cells[1].Range.Text.splitlines()
                    self.text2=self.text1[0]+self.text1[1]
                    print(self.text2)
                    self.wt.Cells(self.row,1).Value=self.text1[0]
                    self.wt.Cells(self.row,2).Value=self.text1[1]
                    # 往excel表格写用例
                    '''for j in range(4,self.length):
                        self.row1=self.row+j-3
                        self.wt.Cells(self.row1,1).Value="step"+str(j-3)
                        self.wt.Cells(self.row1,2).Value=self_word.Tables[i].Rows[j].Cells[1].Range.Text[:-1].strip()
                        self.wt.Cells(self.row1,3).Value=self_word.Tables[i].Rows[j].Cells[2].Range.Text[:-1].strip()'''
                    self.row=self.row+1
                # 计算总用例数
                    self.number_case=self.number_case+1
        except Exception as err:
            print("An exception happend:"+str(err))
            print("报警信息：用例格式有误")
        print("用例个数：",str(self.number_case))
        print("用例总步数：",str(self.number_step))
        self.path=os.path.join(os.getcwd(),"case")
        self.wt.Range("A1:B300").Select()
        self.xlApp.Selection.RowHeight=20
        self.xlApp.Selection.ColumnWidth=25
        self.xlApp.Selection.Font.Name="Arial"
        self.xlApp.Selection.Font.Size=10
        self.xlApp.selection.WrapText = True
        self.xlBook.SaveAs(self.path)
        self.xlBook.Close()
        self.xlApp.Quit()
        self.ms_word.Quit()
    def extract_trac(self,vat_path): # 导出追踪关系函数
        self.xlApp=Dispatch('Excel.Application')
        self.xlBook=self.xlApp.Workbooks.Add()
        self_word=self.lanch_word_1()
        self.vat_path=vat_path
        self.k=0
        self.m=0
        self.number_tables=len(self_word.tables)
        self.number_case=0
        self.wt=self.xlBook.Worksheets("sheet1")
        try:
            for i in range (self.m,self.number_tables):
                # 判断该表格是否为测试用例表格
                if "Case" in self_word.Tables[i].Rows[0].Cells[0].Range.Text:
                    #提取测试用例表格后用例编号的表格
                    self.text=self_word.Tables[i].Rows[0].Cells[1].Range.Text.strip()
                    print(self_word.Tables[i].Rows[0].Cells[1].Range.Text)
                    self.text1=self.text.splitlines()
                    self.length=len(self.text1)
                    self.k=self.k+1
                    self.wt.Cells(self.k,1).Value=self.text1[0]
                    print(self.text1[0])
                    print(self.text1[1])
                    self.wt.Cells(self.k,2).Value=self.text1[1]+self.text1[2]
                    self.wt.Cells(self.k,3).Value=self.text1[3].replace("[Source:","").replace("[[","[").replace("]]","]")
                    self.j=4
                    while self.j<=(self.length-1):
                        if self.text1[self.j] != "" and 'Source'in self.text1[self.j]:
                            self.k=self.k+1
                            self.wt.Cells(self.k,1).Value=self.text1[0].replace("[Source:","").replace("[[","[").replace("]]","]").strip()
                            self.wt.Cells(self.k,2).Value=self.text1[1]+self.text1[2].replace("[Source:","").replace("[[","[").replace("]]","]").strip()
                            self.wt.Cells(self.k,3).Value=self.text1[self.j].replace("[Source:","").replace("[[","[").replace("]]","]")
                            self.wt.Cells(self.k,3).Value=self.wt.Cells(self.k,3).Value.strip()
                        self.j=self.j+1
        except Exception as err:
            print("An exception happend:"+str(err))
            print("报警信息：用例格式有误")
        self.wt.Range("C1:C500").Select()
        self.xlApp.Selection.Replace(" ","")
        # 在VAT中查找需求内容
        self.VAT = self.xlApp.Workbooks.Open(self.vat_path,False)
        self.VAT_sheet=self.VAT.Worksheets(1)
        for i in range(1,self.k):
            for j in range(1,250):
                if self.wt.Cells(i,3).Value != None:
                    if str(self.wt.Cells(i,3).Value).replace("[","").replace("]","") == str(self.VAT_sheet.Cells(j,1).Value).replace("[","").replace("]",""):
                        print(str(self.wt.Cells(i, 3).Value))
                        self.wt.Cells(i,4).Value = self.VAT_sheet.Cells(j,2).Value
                        self.VAT_sheet.Cells(j,1).Select()
                        self.xlApp.Selection.Interior.Color = 5287936
                        break
        self.path=os.path.join(os.getcwd(),"traceability")
        self.wt.Activate()
        self.wt.Range("A1:D300").Select()
        self.xlApp.Selection.RowHeight=20
        self.xlApp.Selection.ColumnWidth=25
        self.xlApp.Selection.Font.Name="Arial"
        self.xlApp.Selection.Font.Size=10
        self.xlApp.selection.WrapText = True
        self.xlBook.SaveAs(self.path)
        self.xlBook.Close()
        self.xlApp.Quit()
        self.ms_word.Quit()
    def extra_req(self):
        try:
            self.word_2=self.lanch_word_1()
            self.path=os.getcwd()
            self.txtName = "req.txt" #将word存到txt文档中
            self.Path=os.path.join(self.path,self.txtName)
            self.file_1=open(self.Path,'w+',encoding='utf-8')
            for para in self.word_2.Paragraphs:
                self.file_1.write(para.Range.Text)
            self.file_1.close()
            self.file=open(self.Path,'r+',encoding='utf-8')
            self.Templine=self.file.readlines()
            print("请输入查找编号")
            self.keyvalue=input()
            self.xlApp = Dispatch('Excel.Application')
            self.xlBook = self.xlApp.Workbooks.Add()
            self.wt=self.xlBook.Worksheets("sheet1")
            self.length=len(self.Templine)
            self.line=1
            self.k=1
            print(self.length)
            while self.line < self.length:
                print(self.Templine[self.line])
                if self.keyvalue in self.Templine[self.line]:
                    self.base_num=0
                    while (self.base_num<=50):
                        if '[End]'in self.Templine[(self.line+self.base_num)]:
                            self.wt.Cells(self.k,1).Value=str(self.Templine[self.line])
                            self.wt.Cells(self.k,2).Value=''.join(self.Templine[self.line:(self.line+self.base_num+1)]).rstrip('\n') 
                            break
                        self.base_num=self.base_num+1
                    self.line=self.line+self.base_num
                    self.k=self.k+1
                self.line=self.line+1
            self.path=os.path.join(os.getcwd(),"extra_req") #需求列表提取到的
            self.xlBook.SaveAs(self.path)
            self.xlBook.Close()
            self.file.close()
            os.unlink(self.Path)
            self.xlApp.Quit()
            self.ms_word.Quit()
        except Exception as err:
            print("An exception happend:"+str(err))
            print("报警信息：文档格式异常")

