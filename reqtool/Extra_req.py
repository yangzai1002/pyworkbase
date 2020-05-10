from PyQt5.QtWidgets import *
import sys, os, shutil
from time import ctime
import time
from PyQt5.QtGui import *
from PyQt5 import QtCore, QtGui
from PyQt5.QtCore import pyqtSignal,pyqtSlot,QObject,QThread
from WordApplication import *
from ExcelApplication import *
from xlwtUtils import *
class ReqExtraWidget(QWidget):

    def __init__(self, parent=None):
        super(ReqExtraWidget,self).__init__(parent)

        self.Reqpath = QLineEdit()
        self.Reqpath.setObjectName("")
        self.Reqpath.setText("")

        self.ReqLable = QLineEdit()
        self.ReqLable.setText("ITCS_RBC_L-SDMS-SwRS")
        self.LogTextEdit = QTextEdit(self)

        #提取需求标签中的特殊标志
        self.Source=QLineEdit()
        self.SourceLabel=QLabel("Source")

        #提取需求标签中的功能特殊标志
        self.function=QLineEdit()
        self.functionLabel=QLabel("Function")

        self.ChooseButton = QPushButton()
        self.ChooseButton.setText("Browse")

        self.ConformButton = QPushButton()
        self.ConformButton.setText("确认")

        self.PathName = QLabel("路径选取")
        self.LableName = QLabel("标签设置")
        self.LogName = QLabel("运行状况")

        layout = QGridLayout()

        layout.addWidget(self.PathName, 0, 0)
        layout.addWidget(self.Reqpath, 0, 1)
        layout.addWidget(self.ChooseButton, 0, 2)

        layout.addWidget(self.LableName, 1, 0)
        layout.addWidget(self.ReqLable, 1, 1)

        #增加源过滤
        layout.addWidget(self.Source,2,1)
        layout.addWidget(self.SourceLabel,2,0);
        #增加功能过滤
        layout.addWidget(self.function,3,1)
        layout.addWidget(self.functionLabel,3,0)
        #增加确认按钮
        layout.addWidget(self.ConformButton, 4, 1)

        layout.addWidget(self.LogName, 5, 0)
        layout_log = layout.addWidget(self.LogTextEdit, 6, 0, 4, 7)

        self.setLayout(layout)

        self.setGeometry(350, 350, 450, 350)
        self.init_UI()
    def init_UI(self):
        #选择路径
        self.ChooseButton.clicked.connect(self.ChooseButton_click)
        #声明一个提取需求类
        self.extrareq = extraReq()
        #链接信号和槽
        self.extrareq.updateLog.connect(self.addlog)
        #链接主程序和信号
        self.ConformButton.clicked.connect(self.extraAct)
    def ChooseButton_click(self):
        # absolute_path is a QString object
        absolute_path, filetype = QFileDialog.getOpenFileName(self, 'Open file',
                                                    '.', "All files (*.*)")
        #设置文本框路径
        self.Reqpath.setText(absolute_path)
    def extraAct(self):
        reply = QMessageBox.question(self, '消息', "确认提取", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.extrareq.setPath(self.Reqpath.text(),
                                 self.ReqLable.text(), self.Source.text()
                                 , self.function.text())
            self.ConformButton.setEnabled(False)
            self.extrareq.finished.connect(self.setConformButtonEnable)
            self.extrareq.start()
        else:
            pass
    #槽函数，发射信息
    def addlog(self,msg):
        self.LogTextEdit.append(msg)
    def setConformButtonEnable(self):
        self.ConformButton.setEnabled(True)
class extraReq(QThread):
    updateLog = pyqtSignal(str)
    def __init__(self):
        super(extraReq,self).__init__()
    def setPath(self,reqpath,reqkey,reqSource,reqFunction):
        self.reqpath=reqpath
        self.reqkey=reqkey
        self.reqSource=reqSource
        self.reqFunction=reqFunction
        #定义输出目录
        self.outpath = os.path.join(os.path.dirname(self.reqpath), "Requirement_specification.xls").replace('/', '\\')
    def sendlog(self,msg):
        self.updateLog.emit("ProcessLog:"+msg)
    def run(self):
        # 输入参数判断
        #1 先判断输入的需求路径
        self.start=time.clock()
        if(os.path.exists(self.reqpath)):
            pass
        else:
            self.sendlog(self.reqpath+"路径不存在,结束处理")
            return
        #2 再判断输入标签是否合理
        if(self.reqkey==""):
            self.sendlog("需求标签为空，结束处理")
            return
        #3 再判断输出路径是否合理
        if(os.path.exists(self.outpath)):
            try:
                os.unlink(self.outpath)
            except Exception as err:
                errorStr = "An exception happend:" + str(err)
                self.sendlog(errorStr)
                return
        #根据输入需求目录生成临时处理文件
        #1 根据需求目录得到临时目录的路径
        self.newReqPath=os.path.join(os.path.dirname(self.reqpath), "temp.doc").replace('/', '\\')
        #2 copy文件，调用函数
        try:
            if(os.path.exists(self.newReqPath)):
                os.unlink(self.newReqPath)
            self.copyFile(self.reqpath,self.newReqPath)
        except Exception as err:
            errorStr = "An exception happend:" + str(err)
            self.sendlog(errorStr)
            return
        self.sendlog("复制新文件:" + self.newReqPath)
        #将word文档提取到临时txt文件中
        self.tempTxtPath = os.path.join(os.path.dirname(self.newReqPath), "req.txt")
        self.wordToTxt(self.newReqPath,self.tempTxtPath)
        self.sendlog("提取txt:" + self.tempTxtPath)
        #将txt文件中的字符读入到列表中
        try:
            file = open(self.tempTxtPath, 'r+', encoding='utf-8')
        except Exception as err:
            errorStr = "An exception happend:" + str(err)
            self.sendlog(errorStr)
        os.chdir(os.path.dirname(self.reqpath))
        #创建一个excel表格
        try:
            myexcel=xlwtUtil()
            myexcel.CreateSheet()
        except Exception as err:
            errorStr = "An exception happend:" + str(err)
            self.sendlog(errorStr)
            return
        #初始化查找参数
        #1 需求标签号
        key = '[' + self.reqkey  # C4D-I-SyRS
        #1 表格行数计数
        sheet_cloumn = 0
        #需求个数计数
        req_num = 0
        try:
            self.sendlog("开始读取" + self.tempTxtPath)
            while True:
                tempLine = file.readline()
                if tempLine == "":
                    break
                if key in tempLine and self.hasChar(tempLine):
                    # 查找的行数递增变量
                    base_num = 1  # TIS-KA_LPS-SwAD,[Reused]
                    # 增加find sourec和find function标志
                    findSource = False
                    findfunction = False
                    isfindSoure = True
                    isfindFunction = True
                    if self.reqSource == "":
                        isfindSoure = False
                    if self.reqFunction == "":
                        isfindFunction = False
                    # 找到标签行
                    reqLabel = str(tempLine)
                    reqLabel = self.defSpace(reqLabel).strip('\n').strip(' ')
                    #设置内容为空，核心查找逻辑处理
                    reqContent = tempLine
                    while (base_num <= 200):
                        # 查找[End]标志
                        tempLine=file.readline()
                        if tempLine == "":
                            break
                        reqContent=reqContent+tempLine
                        if '[End]' in tempLine:
                            findEnd = True
                            break
                        if isfindSoure == True and self.reqSource in tempLine:
                            findSource = True
                        if isfindFunction == True and self.reqFunction in tempLine:
                            findfunction = True
                        if key in tempLine and self.hasChar(tempLine):
                            findEnd = False
                            break
                        # 需求数增加1
                        base_num = base_num + 1
                    # 确定需求内容
                    reqContent=reqContent.strip(' ')
                    # 是否提取需求标签和内容
                    takeSoureLabel = False
                    takeFunctionLabel = False
                    if ((isfindSoure == True and findSource == True) or isfindSoure == False):
                        takeSoureLabel = True
                    if ((isfindFunction == True and findfunction == True) or isfindFunction == False):
                        takeFunctionLabel = True
                    if (takeSoureLabel == True and isfindSoure == True):
                        myexcel.writeCell(sheet_cloumn,2,self.reqSource)
                    if (takeFunctionLabel == True and isfindFunction == True):
                        myexcel.writeCell(sheet_cloumn, 2, ''.join(self.reqFunction))
                    myexcel.writeCell(sheet_cloumn,0,reqLabel)
                    myexcel.writeCell(sheet_cloumn, 1, reqContent)
                    self.sendlog("提取需求"+reqLabel)
                    req_num = req_num + 1
                    sheet_cloumn = sheet_cloumn + 1

            #excel.unionFormat(wt, "A1:C1000")
            myexcel.setWidthAndHeight()
            time.sleep(3)
            #保存excel输出文档
            myexcel.Save(self.outpath)
            self.sendlog("Congratulation,Complete!")
            self.sendlog("文件保存至:" + self.outpath)
            result = "The number of requirement is " + str(req_num)
            self.sendlog(result)
            self.end = time.clock()
            self.sendlog("run time:" + str(self.end-self.start))
        except Exception as err:
            errorStr = "An exception happend:" + str(err)
            self.sendlog(errorStr)
            return
        finally:
            file.close()
            os.unlink(self.newReqPath)
            # 删除临时txt文件
            os.unlink(self.tempTxtPath)
    def hasChar(self, str):
        for i in str:
            if u'\u4e00' <= i <= u'\u9fff':
                return False
            else:
                return True

    def defSpace(self, str):
        for i in str:
            if i == ' ':
                del i
        return str
    def copyFile(self,oldPath,newPath):
        try:
            if os.path.exists(newPath):
                os.unlink(newPath)
            shutil.copy(oldPath, newPath)
        except Exception as err:
            self.sendlog("An exception happend:" + str(err))
            return
    def wordToTxt(self,wordPath,txtPath):
        #创建一个word对象
        try:
            self.word = myWord(wordPath)
            # 接收word文件中的修订记录
            self.word.AcceptRevision()
            # 清楚word中的格式
            self.word.ClearFormat()
            # 删除word中的表格
            self.word.delTable()
        except Exception as err:
            self.sendlog("An exception happend:" + str(err))
            self.word.Save()
         # ITCS_RBC_L-SDMS-SwRS
        try:
            file = open(txtPath, 'w+', encoding='utf-8')
        except Exception as err:
            self.sendlog("An exception happend:" + str(err))
            file.close()
        #获取word中的段落对象
        try:
            Paragraph = self.word.getPara()
            #将word中的段落写到txt文件中去
            for para in Paragraph:
                file.write(para.Range.Text)  # [ITCS_RBC_L-SyRS
        except Exception as err:
            self.sendlog("An exception happend:" + str(err))
        finally:
            self.word.Save()
            self.word.Close()
