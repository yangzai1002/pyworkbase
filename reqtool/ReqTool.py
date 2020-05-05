from Extra_req import *
from Com_req import *
from Trace import *
from log import *
class TabWidget(QTabWidget):
    def __init__(self, parent=None):
        super(TabWidget, self).__init__(parent)
        self.resize(450, 350)
        self.mContent = ReqExtraWidget()
        self.mIndex = IndexWidget()
        self.Trace = TraceWidget()
        self.addTab(self.mContent, u"Extract Requirement")
        self.addTab(self.mIndex, u"Requirement Compare")
        self.addTab(self.Trace, u"Case Trace")
        self.setWindowTitle("Extract")
if __name__ == '__main__':
    import sys
    global work_path
    work_path = "E:\\"
    app = QApplication(sys.argv)
    t=TabWidget()
    app.setWindowIcon(QIcon('2.png'))
    t.show()
    sys.exit(app.exec_())
