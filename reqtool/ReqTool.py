from Extra_req import *
#from Com_req import *
from extraComplete import CompleteWidget
from Trace import *
class TabWidget(QTabWidget):
    def __init__(self, parent=None):
        super(TabWidget, self).__init__(parent)
        self.resize(550, 450)
        self.mContent = ReqExtraWidget()
        #self.mIndex = IndexWidget()
        self.Complete=CompleteWidget()
        self.Trace = TraceWidget()
        self.addTab(self.mContent, u"需求提取")
        #self.addTab(self.mIndex, u"Requirement Compare")
        self.addTab(self.Trace, u"追踪关系提取")
        self.addTab(self.Complete,u"完整性关系提取")
        self.setWindowTitle("需求工作")
if __name__ == '__main__':
    import sys
    global work_path
    work_path = "E:\\"
    app = QApplication(sys.argv)
    t=TabWidget()
    app.setWindowIcon(QIcon('dargon.jpg'))
    t.show()
    sys.exit(app.exec_())
