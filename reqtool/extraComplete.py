from Extra_req import *
class CompleteWidget(QWidget):
    def __init__(self, parent=None):
        super(CompleteWidget, self).__init__(parent)

        self.le = QLineEdit()
        self.le.setObjectName("")
        self.le.setText("")

        self.le1 = QLineEdit()
        self.le1.setText("")

        self.le2 = QLineEdit()
        self.le2.setText("ZDK-SyRTC")

        self.textEdit1 = QTextEdit(self)

        self.pb = QPushButton()
        self.pb.setText("Browse")

        self.pb1 = QPushButton()
        self.pb1.setText("确认")

        self.pb2 = QPushButton()
        self.pb2.setText("Browse")

        self.name1 = QLabel("测试用例")
        self.name2 = QLabel("VAT分配表")
        self.name3 = QLabel("运行状况")
        self.name4 = QLabel("用例标签")
        layout = QGridLayout()

        layout.addWidget(self.name1, 0, 0)
        layout.addWidget(self.le, 0, 1)

        layout.addWidget(self.pb, 0, 2)
        layout.addWidget(self.name2, 1, 0)

        layout.addWidget(self.pb2, 1, 2)

        layout.addWidget(self.le1, 1, 1)
        layout.addWidget(self.pb1, 3, 1)

        layout.addWidget(self.name4, 2, 0)
        layout.addWidget(self.le2,2,1)

        layout.addWidget(self.name3, 4, 0)
        layout_log = layout.addWidget(self.textEdit1, 5, 0, 4, 7)


        self.setLayout(layout)

        self.pb1.clicked.connect(self.startRunCaseTrace)
        self.pb.clicked.connect(self.button_click_1)
        self.pb2.clicked.connect(self.button_click_2)

        self.setGeometry(350, 350, 450, 350)

    def button_click_1(self):
        # absolute_path is a QString object
        absolute_path, filetype = QFileDialog.getOpenFileName(self, 'Open file',
                                                              '', "All files (*.*)")
        self.le.setText(absolute_path)

    def button_click_2(self):
        absolute_path_1, filetype_1 = QFileDialog.getOpenFileName(self, 'Open file',
                                                                  '', "All files (*.*)")
        self.le1.setText(absolute_path_1)

    def __del__(self):
        # Restore sys.stdout
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
    def updateLog(self,msg):
        self.textEdit1.append(msg)
    def startRunCaseTrace(self):
        self.casetrace=caseComplete()
        self.casetrace.setPath(self.le.text(),self.le1.text(),self.le2.text())
        self.casetrace.updateLogSignal.connect(self.updateLog)
        self.pb1.setEnabled(False)
        self.casetrace.start()
        self.casetrace.finished.connect(self.setEnable)
    def setEnable(self):
        self.pb1.setEnabled(True)
class caseComplete(QThread):
    updateLogSignal=pyqtSignal(str)
    def __init__(self):
        super(caseComplete,self).__init__()
    def setPath(self,casePath,vatPath,caseLable):
        self.casePath=casePath
        self.vatPath=vatPath
        self.caseLable=caseLable
    def sendMsg(self,msg):
        self.updateLogSignal.emit("processLog:"+msg)
    def run(self): # 导出追踪关系函数
        casepath=self.casePath.replace('/','\\')
        vatpath=self.vatPath.replace('/','\\')
        savePath = os.path.join(os.path.dirname(self.casePath), "Complete.xlsx")
        savePath = savePath.replace('/', '\\')
        try:
            if os.path.exists(savePath):
                os.unlink(savePath)
        except Exception as err:
            self.sendMsg(err)
        if os.path.exists(casepath) and os.path.exists(vatpath):
            self.sendMsg("开始提取完整性关系，请稍后...")
            pass
        else:
            self.sendMsg('Enter path error')
            return
        try:
            tempath=os.path.join(os.path.dirname(casepath), "temp.doc")
            self.deletePath(tempath)
            shutil.copy(casepath,tempath)
        except Exception as err:
            string = "An exception happend:" + str(err)
            return
        try:
            excel=myExcel()
            wb,wt=excel.AddBook()
        except Exception as err:
            string = "An exception happend:" + str(err)
            return
        try:
            word=myWord(tempath)
            word.AcceptRevision()
            word.ClearFormat()
            Table=word.getTables()
        except Exception as err:
            string = "An exception happend:" + str(err)
            word.Save()
            self.sendMsg(string)
            return
        number_tables = len(Table)
        try:
            #提取用例和需求的对应关系，声明三个列表分别存储用例标签，用例内容，和需求标签
            self.CaseLabel=[]
            self.CaseContent=[]
            self.ReqLabel=[]
            for i in range(0, number_tables):
                # 判断该表格是否为测试用例表格
                if ("Case" in Table[i].Rows[0].Cells[0].Range.Text or "用例" in Table[i].Rows[0].Cells[0].Range.Text) and self.caseLable in Table[i].Rows[0].Cells[1].Range.Text:
                    # 提取测试用例表格后用例编号的表格
                    sourceNum=0
                    sourceLabel=[]
                    text = Table[i].Rows[0].Cells[1].Range.Text.strip().splitlines()
                    for content in text:
                        if 'Source' in content or 'source' in content:
                            sourceLabel.append(self.labelReplace(content))
                            sourceNum=sourceNum+1
                    if self.caseLable in text[0]:
                       for i in range(0,sourceNum):
                           self.CaseLabel.append(str(text[0]).strip(' '))
                           self.CaseContent.append("\n".join(text))
                           self.ReqLabel.append(sourceLabel[i])
                           #wt.Cells(cloumn,1).Value=str(text[0]).strip(' ')
                           #wt.Cells(cloumn,2).Value="\n".join(text)
                           #wt.Cells(cloumn,3).Value=sourceLabel[i]
                           #self.sendMsg(str(wt.Cells(cloumn,3).Value))
                           #cloumn=cloumn+1
            word.Save()
            word.Close()
            os.unlink(tempath)
            vatwb,vatWorksheets=excel.getSheets(vatpath)
            #将VAT中的需求和需求内容都存入到字典中
            #需求字典
            reqDict={}
            for vatws in vatWorksheets:
                vatRow = self.getUsedRow(vatws)
                for m in range(1,vatRow):
                    reqDict[self.labelReplace(vatws.Cells(m,1).Value)]=vatws.Cells(m,2).Value
            vatwb.SaveAs(vatpath)
            #更新完整性关系核心处理流程区
            cloumn = 1
            for key in reqDict.keys():
                onlyOne=0
                wt.Cells(cloumn,1).Value=key
                wt.Cells(cloumn,2).Value=reqDict[key]
                index = 0
                self.sendMsg("需求追踪"+key)
                for label in self.ReqLabel:
                    if(onlyOne==0 and label == key):
                        wt.Cells(cloumn, 3).Value=self.CaseLabel[index]
                        wt.Cells(cloumn, 4).Value = self.CaseContent[index]
                        onlyOne=1
                        cloumn=cloumn+1
                        continue
                    if(onlyOne ==1 and label == key):
                        wt.Cells(cloumn, 1).Value = key
                        wt.Cells(cloumn, 2).Value = reqDict[key]
                        wt.Cells(cloumn, 3).Value = self.CaseLabel[index]
                        wt.Cells(cloumn, 4).Value = self.CaseContent[index]
                        cloumn = cloumn + 1
                        continue
                    index=index+1
            wt.Activate()
            excel.unionFormat(wt,"A1:D1200")
            wb.SaveAs(savePath)
            self.sendMsg("finshed,文件保存至" + savePath)
            vatwb.SaveAs(vatpath)
            reqDict.clear()
        except Exception as err:
            string = "An exception happend:" + str(err)
            self.sendMsg(string)
    def labelReplace(self,str1):
        if (str1 ==None):
            return
        string=str1.replace("[Source:", "").replace("[[", "[").replace("]]", "]").strip()
        return str(string)
    def getUsedRow(self,ws):
        info = ws.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count
        return nrows
    def deletePath(self,path):
        if os.path.exists(path):
            os.unlink(path)