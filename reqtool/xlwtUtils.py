import xlwt,os
import time
class xlwtUtil():
    def __init__(self):
        pass

    def CreateSheet(self):
        # 创建一个workbook 设置编码
        self.workbook = xlwt.Workbook(encoding='utf-8')
        # 创建一个worksheet
        self.worksheet = self.workbook.add_sheet('My Worksheet')
        self.font = self.GetFont()
    def writeCell(self,row,cloumn,value):
        self.worksheet.write(row, cloumn, value, self.font)
    def getCellValue(self):
        pass
    def Save(self,path):
        self.workbook.save(path)
    def setWidthAndHeight(self):
        tall_style = xlwt.easyxf('font:height 230')  # 36pt
        for i in range(0,10):
            self.worksheet.col(i).width = 35 * 256
        for i in range(0,2000):
            self.worksheet.row(i).height_mismatch = True
            self.worksheet.row(i).height = 20 * 35  # 20为基准数，40意为40磅
    def GetFont(self):
        style = xlwt.XFStyle()  # 初始化样式
        font = xlwt.Font()  # 为样式创建字体
        font.name = "微软雅黑"
        font.bold = False  # 黑体
        font.underline = False  # 下划线
        font.italic = False  # 斜体字
        font.height = 20 * 12
        style.font = font  # 设定样式

        alignment = xlwt.Alignment()
        # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        alignment.horz = 0x01
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        alignment.vert = 0x00
        # 设置自动换行
        alignment.wrap = 1
        style.alignment=alignment

        # 设置边框
        borders = xlwt.Borders()
        # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
        # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        borders.left_colour = 255
        borders.right_colour = 255
        borders.top_colour = 255
        borders.bottom_colour = 255
        style.borders=borders
        # 设置列宽，一个中文等于两个英文等于两个字符，11为字符数，256为衡量单位
        #

        # 设置背景颜色
        pattern = xlwt.Pattern()
        # 设置背景颜色的模式
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        # 背景颜色
        pattern.pattern_fore_colour = 255

        style.pattern=pattern

        return style
if __name__=="__main__":
    myexcel=xlwtUtil()
    myexcel.CreateSheet()
    str="[TIS-KA-TISPS-SsyRS-0001]TISPS shall be a 2-by-2-out-2 system consisting of one or more chassis, including two redundant 2-out-2 systems.TISPS 应是由一个或多个机箱组成的一个2乘2 取2 系统，包含2组冗余的2取2系统。#Comment=[Reused] [<LKD2-KA-TCPS-SsyRS-0001>,<IPS-200-SyRS-0001>]#Source=[<TIS-KA-SyAD-0492>,<TIS-KA-SyAD-0010>]#Category= Non-Functional#Contribution=RAM[End]"
    start=time.clock()
    for i in range(0,1000):
        myexcel.writeCell(i,0,"zhang晋阳")
        myexcel.writeCell(i,1,str)
        myexcel.setWidthAndHeight()
    path=os.path.join(os.getcwd(),"vat.xls")
    if(os.path.exists(path)):
        os.unlink(path)
    print(path)
    myexcel.Save(path)
    end=time.clock()
    print((end-start))
