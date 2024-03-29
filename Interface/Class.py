from PyQt5.QtWidgets import QMainWindow
from Interface.MS import Ui_MSForm

class MSMainForm(QMainWindow, Ui_MSForm):
    def __init__(self, parent=None):
        super(MSMainForm, self).__init__(parent)
        self.setupUi(self)
        from PyQt5.QtCore import QDate
        self.DateDE.setDate(QDate.currentDate())  # 设置成当天日期
        self.ULD951Btn.clicked.connect(self.ULD951Fctn)  # ULD951功能
        self.LWS952Btn.clicked.connect(self.LWS952Fctn)  # LWS952功能
        self.DSMnfstBtn.clicked.connect(self.DSMnfstFctn)  # DS舱单功能
        self.ULDMnfstBtn.clicked.connect(self.ULDMnfstFctn)  # ULD舱单功能

    def ULD951Fctn(self):  # ULD951功能
        self.MsgLabel.setText("ULD951运行中")
        self.MsgLabel.repaint()  # MsgLabel重绘
        Year2Month2 = self.DateDE.date().toString('yyMM')  # 2位数字年 + 2位数字月
        ULD951DirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/财务/CPM UCM/'  # CPM951目录路径
        from PyQt5.QtCore import QLocale
        Day2MonthEA = QLocale(QLocale.English).toString(self.DateDE.date(), 'ddMMM').upper()  # 2位数字日 + 大写英语缩写月
        CPM951FileName = 'CPM MS951-' + Day2MonthEA + '.pdf'  # CPM951文件名
        CPM951FilePath = ULD951DirPath + CPM951FileName  # CPM951文件路径
        import os
        if os.path.exists(CPM951FilePath) == False:  # CPM951文件不存在
            self.MsgLabel.setText("CPM951文件不存在！！")
            return
        AKE951FilePath = 'C:/Files/MS/ULD/AKE MS951.xlsx'
        if os.path.exists(AKE951FilePath) == False:  # AKE951文件不存在
            self.MsgLabel.setText("AKE951文件不存在！！")
            return
        ULDStkDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/3/'  # ULDStock目录路径
        import datetime
        DateDT = datetime.datetime.strptime(Day2MonthEA, '%d%b')  # 日期字符串转日期格式
        DateDT = DateDT + datetime.timedelta(days=1)  # 日期加1天
        Date = DateDT.strftime('%d%b').upper()  # 日期格式转日期字符串并大写
        ULDStkFileName = Date + ' PVG ULD STOCK.xlsx'  # ULDStock文件名
        ULDStkFilePath = ULDStkDirPath + ULDStkFileName  # ULDStock文件路径
        if os.path.exists(ULDStkFilePath) == False:  # ULDStock文件不存在
            self.MsgLabel.setText("ULDStock文件不存在！！")
            return
        from ReadPDF.Function import ReadPage1PDF
        CPM951 = ReadPage1PDF(CPM951FilePath)  # 返回CPM951第1页PDF文件文本
        from ReadPDF.CPM.Function import ReadCPM
        ReadCPM(CPM951)  # 读取CPM951
        Year2Month2Day2 = self.DateDE.date().toString('yyMMdd')  # 2位数字年 + 2位数字月 + 2位数字日
        from WritExcl.AKE951.Function import WritAKE951ST
        WritAKE951ST(AKE951FilePath, Year2Month2Day2)  # 写AKE951页
        from WritExcl.ULDStock.Function import WritCPMULDStkST
        WritCPMULDStkST(ULDStkFilePath)  # 写ULDStock页
        UCM951FileName = 'UCM MS951-' + Day2MonthEA + '.pdf'  # UCM951文件名
        UCM951FilePath = ULD951DirPath + UCM951FileName  # UCM951文件路径
        if os.path.exists(UCM951FilePath) == False:  # UCM951文件不存在
            self.MsgLabel.setText("UCM951文件不存在！！")
            return
        UCM951 = ReadPage1PDF(UCM951FilePath)  # 返回UCM951第1页PDF文件文本
        from ReadPDF.UCM.Function import ReadUCM
        ReadUCM(UCM951)  # 读取UCM
        from WritExcl.ULDStock.Function import WritUCMULDStkST
        WritUCMULDStkST(ULDStkFilePath)  # 写UCM到ULDStock页
        self.MsgLabel.setText("ULD951录入完成")

    def LWS952Fctn(self):  # LWS952功能
        self.MsgLabel.setText("LWS952运行中")
        Year2Month2 = self.DateDE.date().toString('yyMM')  # 2位数字年 + 2位数字月
        Day2 = self.DateDE.date().toString('dd')  # 2位数字日
        OutDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/航班/' + Day2 + '/OUT/'  # OUT目录路径
        from PyQt5.QtCore import QLocale
        Day2MonthEA = QLocale(QLocale.English).toString(self.DateDE.date(), 'ddMMM').upper()  # 2位数字日 + 大写英语缩写月
        LWS952FileName = 'DLWS ' + Day2MonthEA + '.pdf'  # LWS952文件名
        LWS952FilePath = OutDirPath + LWS952FileName  # LWS952文件路径
        import os
        if os.path.exists(LWS952FilePath) == False:  # LWS952文件不存在
            self.MsgLabel.setText("LWS952文件不存在！！")
            return
        ULDStkDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/3/'  # ULDStock目录路径
        ULDStkFileName = Day2MonthEA + ' PVG ULD STOCK.xlsx'  # ULDStock文件名
        ULDStkFilePath = ULDStkDirPath + ULDStkFileName  # ULDStock文件路径
        if os.path.exists(ULDStkFilePath) == False:  # ULDStock文件不存在
            self.MsgLabel.setText("ULDStock文件不存在！！")
            return
        from ReadPDF.Function import ReadPage1PDF
        LWS952 = ReadPage1PDF(LWS952FilePath)  # 返回LWS952第1页PDF文件文本
        from ReadPDF.LWS.Function import ReadLWS
        ReadLWS(LWS952)  # 读取LWS
        from WritExcl.ULDStock.Function import DelLWSULDStkST
        DelLWSULDStkST(ULDStkFilePath)  # 删除LWS集装器在ULD Stock页
        self.MsgLabel.setText("LWS952更新完成")

    def DSMnfstFctn(self):  # DS舱单功能
        self.MsgLabel.setText("DS舱单运行中")
        self.MsgLabel.repaint()  # MsgLabel重绘
        Year2Month2 = self.DateDE.date().toString('yyMM')  # 2位数字年 + 2位数字月
        Day2 = self.DateDE.date().toString('dd')  # 2位数字日
        OutDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/航班/' + Day2 + '/OUT/'  # OUT目录路径
        MnfstFilePath = OutDirPath + '舱单 - 副本.xlsx'  # 舱单副本表格文件路径
        import os
        if os.path.exists(MnfstFilePath) == False:  # 舱单副本表格文件不存在
            self.MsgLabel.setText("舱单副本表格文件不存在！！")
            return
        from ReadExcl.Mnfst.Function import ReadFltMnfstST
        ReadFltMnfstST(MnfstFilePath)  # 读取舱单副本表格航班舱单页
        from ReadExcl.Mnfst.Function import ReadULDMnfstST
        ReadULDMnfstST(MnfstFilePath)  # 读取舱单副本表格ULD舱单页
        from WritExcl.DSMnfst.Function import ReadMnfstLst
        ReadMnfstLst()  # 读取舱单对象列表
        from WritExcl.DSMnfst.Function import WritDSMnfstST
        WritDSMnfstST(MnfstFilePath)  # 写DS舱单页
        self.MsgLabel.setText("DS舱单录入完成")

    def ULDMnfstFctn(self):  # ULD舱单功能
        self.MsgLabel.setText("ULD舱单运行中")
        self.MsgLabel.repaint()  # MsgLabel重绘
        Year2Month2 = self.DateDE.date().toString('yyMM')  # 2位数字年 + 2位数字月
        Day2 = self.DateDE.date().toString('dd')  # 2位数字日
        OutDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/航班/' + Day2 + '/OUT/'  # OUT目录路径
        MnfstFilePath = OutDirPath + '舱单 - 副本.xlsx'  # 舱单副本表格文件路径
        import os
        if os.path.exists(MnfstFilePath) == False:  # 舱单副本表格文件不存在
            self.MsgLabel.setText("舱单副本表格文件不存在！！")
            return
        Year4_Month2_Day2 = self.DateDE.date().toString('yyyy-MM-dd')  # 4位数字年-2位数字月-2位数字日
        ULDMnfstFileName = '_lkg_gsa_ffm_舱单_' + Year4_Month2_Day2 + '.xlsx'  # ULD舱单表格文件名
        ULDMnfstFilePath = OutDirPath + ULDMnfstFileName  # ULD舱单表格文件路径
        if os.path.exists(ULDMnfstFilePath) == False:  # ULD舱单表格文件不存在
            self.MsgLabel.setText("ULD舱单表格文件不存在！！")
            return
        from ReadExcl.Mnfst.Function import ReadFltMnfstST
        ReadFltMnfstST(MnfstFilePath)  # 读取舱单副本表格航班舱单页
        from ReadExcl.Mnfst.Function import ReadULDMnfstST
        ReadULDMnfstST(MnfstFilePath)  # 读取舱单副本表格ULD舱单页
        from WritExcl.ULDMnfst.Function import WritULDMnfstST
        WritULDMnfstST(ULDMnfstFilePath)  # 写集装器舱单信息页
        self.MsgLabel.setText("ULD舱单录入完成")