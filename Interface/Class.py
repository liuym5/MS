from PyQt5.QtWidgets import QMainWindow
from Interface.MS import Ui_MSForm

class MSMainForm(QMainWindow, Ui_MSForm):
    def __init__(self, parent=None):
        super(MSMainForm, self).__init__(parent)
        self.setupUi(self)
        from PyQt5.QtCore import QDate
        self.DateDE.setDate(QDate.currentDate())  # 设置成当天日期
        self.ULDMnfstBtn.clicked.connect(self.ULDMnfstFctn)  # ULD舱单功能

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
        ULDMnfstPath = OutDirPath + '_lkg_gsa_ffm_舱单_2024-1-19.xlsx'  # ULD舱单表格文件路径
        if os.path.exists(ULDMnfstPath) == False:  # ULD舱单表格文件不存在
            self.MsgLabel.setText("ULD舱单表格文件不存在！！")
            return
        from ReadExcl.Mnfst.Function import ReadFltMnfstST
        ReadFltMnfstST(MnfstPath)  # 读取舱单副本表格航班舱单页
        from ReadExcl.Mnfst.Function import ReadULDMnfstST
        ReadULDMnfstST(MnfstPath)  # 读取舱单副本表格ULD舱单页
        self.MsgLabel.setText("ULD舱单录入完成")