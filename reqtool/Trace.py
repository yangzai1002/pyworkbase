﻿from Extra_req import *
class TraceWidget(QWidget):
    def __init__(self, parent=None):
        super(TraceWidget, self).__init__(parent)

        self.le = QLineEdit()
        self.le.setObjectName("")
        self.le.setText("")

        self.le1 = QLineEdit()
        self.le1.setText("")

        self.le2 = QLineEdit()
        self.le2.setText("ITCS_RBC_L-SDMS-SwRTC")

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
        casetrace=caseTrace()
        casetrace.setPath(self.le.text(),self.le1.text(),self.le2.text())
        casetrace.updateLogSignal.connect(self.updateLog())
        casetrace.start()
class caseTrace(QThread):
    updateLogSignal=pyqtSignal(QObject)
    def __init__(self):
        super(caseTrace,self).__init__()
    def setPath(self,casePath,vatPath,caseLable):
        self.casePath=casePath
        self.vatPath=vatPath
        self.caseLable=caseLable
    def sendMsg(self,msg):
        self.updateLogSignal.emit(msg)
    def case_trace(self): # 导出追踪关系函数
        casepath=self.casePath.replace('/','\\')
        vatpath=self.vatPath.replace('/','\\')
        if os.path.exists(casepath) and os.path.exists(vatpath):
            self.sendMsg("开始提取追踪关系，请稍后...")
            pass
        else:
            self.sendMsg('Enter path error')
            return
        tempath=os.path.join(os.path.dirname(casepath), "temp.doc")
        self.deletePath(tempath)
        shutil.copy(casepath,tempath)
        excel=myExcel()
        wb,wt=excel.AddBook()
        word=myWord(tempath)
        word.AcceptRevision()
        word.ClearFormat()
        Table=word.getTables()
        cloumn = 1
        number_tables = len(Table)
        try:
            for i in range(0, number_tables):
                # 判断该表格是否为测试用例表格
                if ("Case" in Table[i].Rows[0].Cells[0].Range.Text or "用例" in Table[i].Rows[0].Cells[0].Range.Text) and caseLable in Table[i].Rows[0].Cells[1].Range.Text:
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
                           wt.Cells(cloumn,1).Value=str(text[0]).strip(' ')
                           wt.Cells(cloumn,2).Value="\n".join(text)
                           wt.Cells(cloumn,3).Value=sourceLabel[i]
                           self.sendMsg(str(wt.Cells(cloumn,3).Value))
                           cloumn=cloumn+1
            word.Save()
            word.Close()
            os.unlink(tempath)
            vatwb,vatWorksheets=excel.getSheets(vatpath)
            traceCloumn=self.getUsedRow(wt)
            for vatws in vatWorksheets:
                vatCloumn=self.getUsedRow(vatws)
                for m in range(1, traceCloumn+1):
                    if wt.Cells(m, 3).Value != None:
                        for n in range(1, vatCloumn+1):
                            if vatws.Cells(n, 1).Value != None:
                                if self.labelReplace(wt.Cells(m, 3).Value) == self.labelReplace(vatws.Cells(n, 1).Value):
                                    self.sendMsg(str(wt.Cells(m, 3).Value))
                                    wt.Cells(m, 4).Value = vatws.Cells(n, 2).Value
                                    vatws.Cells(n, 1).Interior.Color = 5287936
                                    break
            savePath=os.path.join(os.path.dirname(self.casePath), "traceability.xlsx")
            savePath=savePath.replace('/','\\')
            wt.Activate()
            excel.unionFormat(wt,"A1:D800")
            if os.path.exists(savePath):
                os.unlink(savePath)
            wb.SaveAs(savePath)
            vatwb.SaveAs(vatpath)
        except Exception as err:
            string="An exception happend:" + str(err)
            self.sendMsg(string)
        finally:
            excel.Quit()
            self.sendMsg("finshed,文件保存至"+savePath)
            word.Quit()
    def labelReplace(self,str1):
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