from PyQt5.QtWidgets import QMainWindow
from Interface.MS import Ui_MSForm

class MSMainForm(QMainWindow, Ui_MSForm):
    def __init__(self, parent=None):
        super(MSMainForm, self).__init__(parent)
        self.setupUi(self)
        from PyQt5.QtCore import QDate
        self.DateDE.setDate(QDate.currentDate())  # 设置成当天日期
        self.DSMnfstBtn.clicked.connect(self.DSMnfstFctn)  # DS舱单功能
        self.ULDMnfstBtn.clicked.connect(self.ULDMnfstFctn)  # ULD舱单功能

    def DSMnfstFctn(self):  # DS舱单功能
        self.MsgLabel.setText("DS舱单运行中")
        self.MsgLabel.repaint()  # MsgLabel重绘
        Year2Month2 = self.DateDE.date().toString('yyMM')  # 2位数字年 + 2位数字月
        Day2 = self.DateDE.date().toString('dd')  # 2位数字日
        OutDirPath = 'C:\Files\MS\日常\\' + Year2Month2 + '\航班\\' + Day2 + '\OUT\\'  # OUT目录路径
        MnfstPath = OutDirPath + '舱单 - 副本.xlsx'  # 舱单副本表格文件路径
        import os
        if os.path.exists(MnfstPath) == False:  # 舱单副本表格文件不存在
            self.MsgLabel.setText("舱单副本表格文件不存在！！")
            return
        from ReadExcl.Mnfst.Function import ReadFltMnfstST
        ReadFltMnfstST(MnfstPath)  # 读取舱单副本表格航班舱单页
        from ReadExcl.Mnfst.Function import ReadULDMnfstST
        ReadULDMnfstST(MnfstPath)  # 读取舱单副本表格ULD舱单页
        from WritExcl.DSMnfst.Function import ReadMnfstLst
        ReadMnfstLst()  # 读取舱单对象列表
        from WritExcl.DSMnfst.Function import WritDSMnfst
        WritDSMnfst(MnfstPath)  # 写DS舱单页
        self.MsgLabel.setText("DS舱单录入完成")

    def ULDMnfstFctn(self):  # ULD舱单功能
        self.MsgLabel.setText("ULD舱单运行中")
        self.MsgLabel.repaint()  # MsgLabel重绘
        Year2Month2 = self.DateDE.date().toString('yyMM')  # 2位数字年 + 2位数字月
        Day2 = self.DateDE.date().toString('dd')  # 2位数字日
        OutDirPath = 'C:\Files\MS\日常\\' + Year2Month2 + '\航班\\' + Day2 + '\OUT\\'  # OUT目录路径
        MnfstPath = OutDirPath + '舱单 - 副本.xlsx'  # 舱单副本表格文件路径
        import os
        if os.path.exists(MnfstPath) == False:  # 舱单副本表格文件不存在
            self.MsgLabel.setText("舱单副本表格文件不存在！！")
            return
        Year4_Month2_Day2 = self.DateDE.date().toString('yyyy-MM-dd')  # 4位数字年-2位数字月-2位数字日
        ULDMnfstFile = '_lkg_gsa_ffm_舱单_' + Year4_Month2_Day2 + '.xlsx'  # ULD舱单表格文件名
        ULDMnfstPath = OutDirPath + ULDMnfstFile  # ULD舱单表格文件路径
        if os.path.exists(ULDMnfstPath) == False:  # ULD舱单表格文件不存在
            self.MsgLabel.setText("ULD舱单表格文件不存在！！")
            return
        from ReadExcl.Mnfst.Function import ReadFltMnfstST
        ReadFltMnfstST(MnfstPath)  # 读取舱单副本表格航班舱单页
        from ReadExcl.Mnfst.Function import ReadULDMnfstST
        ReadULDMnfstST(MnfstPath)  # 读取舱单副本表格ULD舱单页
        from WritExcl.ULDMnfst.Function import WritULDMnfst
        WritULDMnfst(ULDMnfstPath)  # 写集装器舱单信息页
        self.MsgLabel.setText("ULD舱单录入完成")