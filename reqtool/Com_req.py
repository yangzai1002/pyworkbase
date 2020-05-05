from Extra_req import *
from ExcelApplication import *
class IndexWidget(ReqExtraWidget):
    def __init__(self,parent=None):
        super(ReqExtraWidget, self).__init__(parent)


        self.le = QLineEdit()
        self.le.setObjectName("")
        self.le.setText("")

        self.le1 = QLineEdit()
        self.le1.setText("")

        self.label = QLineEdit()
        self.label.setText("ITCS_RBC_L-SDMS-SwRS")

        self.textEdit1 = QTextEdit()

        self.pb = QPushButton()
        # self.pb.setObjectName("browse")
        self.pb.setText("Browse")

        self.pb1 = QPushButton()
        self.pb1.setText("确认")

        self.pb2 = QPushButton()
        # self.pb2.setObjectName("browse")
        self.pb2.setText("Browse")

        self.name1 = QLabel("修改前文档")
        self.name2 = QLabel("修改后文档")
        self.name3 = QLabel("运行状况")
        self.name4 = QLabel("标签设置")

        layout1 = QGridLayout()

        layout1.addWidget(self.name1, 0, 0)
        layout1.addWidget(self.le, 0, 1)

        layout1.addWidget(self.pb, 0, 2)
        layout1.addWidget(self.name2, 1, 0)

        layout1.addWidget(self.pb2, 1, 2)

        layout1.addWidget(self.le1, 1, 1)
        layout1.addWidget(self.pb1, 3, 1)

        layout1.addWidget(self.name4, 2, 0)
        layout1.addWidget(self.label, 2, 1)

        layout1.addWidget(self.name3, 4, 0)
        layout_log = layout1.addWidget(self.textEdit1, 5, 0, 4, 7)

        # layout.setRowStretch(3, 1);
        # layout.setRowStretch(4, 3);

        self.setLayout(layout1)

        self.pb.clicked.connect(self.button_click_1)
        self.pb2.clicked.connect(self.button_click_2)
        self.pb1.clicked.connect(self.ExtraEvent)
        # self.square=QFrame(self)
        # self.square.setGeometry(150,20,100,100)
        self.setGeometry(350, 350, 450, 350)
        # self.setWindowTitle("Extraction")

        self.path1="E:\\requirement1.xlsx"
        self.path2="E:\\requirement2.xlsx"

    def __del__(self):
        # Restore sys.stdout
        sys.stdout_1 = sys.__stdout__
        sys.stderr_1 = sys.__stderr__

    def button_click_1(self):
        # absolute_path is a QString object
        absolute_path, filetype = QFileDialog.getOpenFileName(self, 'Open file',
                                                              '.', "All files (*.*)")
        self.le.setText(absolute_path)

    def button_click_2(self):
        absolute_path_1, filetype_1 = QFileDialog.getOpenFileName(self, 'Open file',
                                                                  '.', "All files (*.*)")
        self.le1.setText(absolute_path_1)


    def ExtraEvent(self, event):
        reply = QMessageBox.question(self, '消息', "确认提取", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.run_compare_req()
        else:
            pass
    def getUsedRow(self,ws):
        info = ws.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count
        return nrows
    def getBig(self,num1,num2):
        if num1>num2:
            return num1
        else:
            return num2
    def giveData(self,ws_com,ws_pre,ws_aft,cloumn,i,j):
        ws_com.Cells(cloumn, 1).Value = ws_pre.Cells(i, 1).Value
        ws_com.Cells(cloumn, 2).Value = ws_pre.Cells(i, 2).Value
        ws_com.Cells(cloumn, 3).Value = ws_aft.Cells(j, 2).Value
    def ispath(self,path1,path2):
        if not os.path.exists(path1) or not os.path.exists(path2):
            return False
        return True
    def Compare_req(self,pathpre,pathaft,labelkey):
        # 获取路径
        log=Logger("compareLog.log")
        try:
            log.logger.info("Start,Waiting...")
            path_pre = pathpre
            path_aft = pathaft
            if not self.ispath(pathpre,pathaft):
                log.logger.info("请输入正确的路径")
                return
            items=[]
            path1 = [path_pre, self.path1]
            path2= [path_aft,self.path2]
            item=[path1,path2]
            path_item=dict(item)
            for key,value in path_item.items():
                super().extra_req(key,labelkey,value)
            excel=myExcel()
            log.logger.info(self.path1)
            wb_pre,ws_pre=excel.OpenBook(self.path1,"sheet1")
            wb_aft,ws_aft=excel.OpenBook(self.path2,"sheet1")
            wb_com,ws_com = excel.AddBook()
            cloumn = 2
            ws_com.Cells(1, 1).Value = "需求编号"
            ws_com.Cells(1, 2).Value = "更新前需求"
            ws_com.Cells(1, 3).Value = "更新后需求"
            log.logger.info("开始对比文档差异，请稍后...")
            row_pre=self.getUsedRow(ws_pre)
            row_aft=self.getUsedRow(ws_aft)
            row=self.getBig(row_pre,row_aft)
            old_req = []
            NumberChangeReq=0
            for i in range(1, row):
                if labelkey in str(ws_pre.Cells(i, 1).Value):
                    req_find = False
                    pre_text = str(ws_pre.Cells(i, 2).Value).splitlines()
                    for j in range(1, row):
                        aft_text = str(ws_aft.Cells(j, 2).Value).splitlines()
                        if str(ws_pre.Cells(i, 1).Value) == str(ws_aft.Cells(j, 1).Value):
                            req_find=True
                            self.giveData(ws_com, ws_pre, ws_aft, cloumn, i, j)
                            old_req.append(j)
                            if pre_text == aft_text:
                                log.logger.info(str(ws_com.Cells(cloumn, 1).Value)+"无变化")
                            else:
                                log.logger.info(str(ws_com.Cells(cloumn, 1).Value)+"有变化")
                                NumberChangeReq=NumberChangeReq+1
                                ws_com.Rows(cloumn).Select()
                                excel.makeColor(65535)
                            cloumn=cloumn+1
                    if req_find == False:
                        ws_com.Cells(cloumn, 1).Value = ws_pre.Cells(i, 1).Value
                        ws_com.Cells(cloumn, 2).Value = ws_pre.Cells(i, 2).Value
                        log.logger.info(str(ws_com.Cells(cloumn, 1).Value) + "需求已删除")
                        NumberChangeReq=NumberChangeReq+1
                        ws_com.Rows(cloumn).Select()
                        excel.makeColor(255)
                        cloumn = cloumn + 1
            for i in range(1, row):
                if labelkey in str(ws_aft.Cells(i, 1).Value):
                    if i not in old_req:
                        ws_com.Cells(cloumn, 2).Value = ws_aft.Cells(i, 1).Value
                        ws_com.Cells(cloumn, 3).Value = ws_aft.Cells(i, 2).Value
                        # print(str(self.ws_compare.Cells(self.cloumn,2).Value))
                        ws_com.Rows(cloumn).Select()
                        excel.makeColor(5287936)
                        NumberChangeReq = NumberChangeReq + 1
                        log.logger.info(str(ws_com.Cells(cloumn, 2).Value) + "新增需求")
                        cloumn = cloumn + 1
            excel.unionFormat(ws_com,"A1:C400")
            ChangePercent=NumberChangeReq/row_pre
            if os.path.exists("E:\\对比报告.xlsx"):
                os.unlink("E:\\对比报告.xlsx")
            wb_com.SaveAs("E:\\对比报告.xlsx")
            log.logger.info("Congratulation,Compare report Complete!")
            log.logger.info("需求变化率为"+str(ChangePercent))
            log.logger.info("对比报告保存至E:\对比报告.xlsx")
        except Exception as err:
            errorStr="An exception happend:" + str(err)
            log.logger.info(errorStr)
        finally:
            excel.Quit()
    def run_compare_req(self):
        sys.stdout = EmittingStream(textWritten=self.normalOutputWritten1)
        sys.stderr = EmittingStream(textWritten=self.normalOutputWritten1)
        thread.start_new_thread(self.Compare_req, (self.le.text(),self.le1.text(),self.label.text()))
    def normalOutputWritten1(self, text):
        """Append text to the QTextEdit."""
        # Maybe QTextEdit.append() works as well, but this is how I do it:
        cursor = self.textEdit1.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.textEdit1.setTextCursor(cursor)
        self.textEdit1.ensureCursorVisible()