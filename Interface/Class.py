from PyQt5.QtWidgets import QMainWindow
from Interface.MS import Ui_MSForm

class MSMainForm(QMainWindow, Ui_MSForm):
    def __init__(self, parent=None):
        super(MSMainForm, self).__init__(parent)
        self.setupUi(self)
        from PyQt5.QtCore import QDate
        self.DateDE.setDate(QDate.currentDate())  # 设置成当天日期
        self.ULD951Btn.clicked.connect(self.ULD951Fctn)  # ULD951功能
        self.PRELOADBtn.clicked.connect(self.PRELOADFctn)  # PRELOAD功能
        self.ULDMnfstBtn.clicked.connect(self.ULDMnfstFctn)  # ULD舱单功能
        self.UCM951Btn.clicked.connect(self.UCM951Fctn)  # UCM951功能
        self.ULD952Btn.clicked.connect(self.ULD952Fctn)  # ULD952功能
        self.LWS952Btn.clicked.connect(self.LWS952Fctn)  # LWS952功能
        self.AKE952Btn.clicked.connect(self.AKE952Fctn)  # AKE952功能
        self.RentalMnfstBtn.clicked.connect(self.RentalMnfstFctn)  # 租板舱单功能
        self.SCM1Btn.clicked.connect(self.SCM1Fctn)  # SCM1功能
        self.SCM2Btn.clicked.connect(self.SCM2Fctn)  # SCM2功能

    def ULD951Fctn(self):  # ULD951功能
        self.MsgLabel.setText("ULD951运行中")
        self.MsgLabel.repaint()  # MsgLabel重绘
        Year2Month2 = self.DateDE.date().toString('yyMM')  # 2位数字年 + 2位数字月
        ULD951DirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/财务/CPM UCM/'  # CPM951目录路径
        from PyQt5.QtCore import QLocale
        Day2MonthEA = QLocale(QLocale.English).toString(self.DateDE.date(), 'ddMMM').upper()  # 2位数字日 + 大写英语缩写月
        CPM951FileName = 'CPM MS951-' + Day2MonthEA + '.txt'  # CPM951文件名
        CPM951FilePath = ULD951DirPath + CPM951FileName  # CPM951文件路径
        import os
        if os.path.exists(CPM951FilePath) == False:  # CPM951文件不存在
            self.MsgLabel.setText("CPM951文件不存在！！")
            return
        AKE951FilePath = 'C:/Files/MS/ULD/AKE MS951.xlsx'
        if os.path.exists(AKE951FilePath) == False:  # AKE951文件不存在
            self.MsgLabel.setText("AKE951文件不存在！！")
            return
        Day2MonthEAYear2 = QLocale(QLocale.English).toString(self.DateDE.date(), 'ddMMMyy').upper()  # 2位数字日 + 大写英语缩写月 + 2位数字年
        import datetime
        DateDT = datetime.datetime.strptime(Day2MonthEAYear2, '%d%b%y')  # 日期字符串转日期格式
        DateDT = DateDT + datetime.timedelta(days=1)  # 日期加1天
        Year2Month2 = DateDT.strftime('%y%m')  # 日期格式转日期字符串并大写
        ULDStkDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/5/'  # ULDStock目录路径
        Date = DateDT.strftime('%d%b').upper()  # 日期格式转日期字符串并大写
        ULDStkFileName = Date + ' PVG ULD STOCK.xlsx'  # ULDStock文件名
        ULDStkFilePath = ULDStkDirPath + ULDStkFileName  # ULDStock文件路径
        if os.path.exists(ULDStkFilePath) == False:  # ULDStock文件不存在
            self.MsgLabel.setText("ULDStock文件不存在！！")
            return
        from ReadTXT.CPM.Function import ReadCPM
        StackTF = ReadCPM(CPM951FilePath)  # 读取CPM951,返回是否有叠板
        Year2Month2Day2 = self.DateDE.date().toString('yyMMdd')  # 2位数字年 + 2位数字月 + 2位数字日
        from WritExcl.AKE951.Function import WritAKE951ST
        WritAKE951ST(AKE951FilePath, Year2Month2Day2)  # 写AKE951页
        from WritExcl.ULDStk.Function import WritCPMULDStkST
        WritCPMULDStkST(ULDStkFilePath)  # 写ULDStock页
        if StackTF:  # 有叠板
            UCM951FileName = 'UCM MS951-' + Day2MonthEA + '.txt'  # UCM951文件名
            UCM951FilePath = ULD951DirPath + UCM951FileName  # UCM951文件路径
            if os.path.exists(UCM951FilePath) == False:  # UCM951文件不存在
                self.MsgLabel.setText("UCM951文件不存在！！")
                return
            from ReadTXT.UCM951.Function import ReadUCM
            ReadUCM(UCM951FilePath)  # 读取UCM
            from WritExcl.ULDStk.Function import WritUCMULDStkST
            WritUCMULDStkST(ULDStkFilePath)  # 写UCM到ULDStock页
        self.MsgLabel.setText("ULD951录入完成")

    def PRELOADFctn(self):  # PRELOAD功能
        self.MsgLabel.setText("PRELOAD运行中")
        self.MsgLabel.repaint()  # MsgLabel重绘
        Year2Month2 = self.DateDE.date().toString('yyMM')  # 2位数字年 + 2位数字月
        Day2 = self.DateDE.date().toString('dd')  # 2位数字日
        OutDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/航班/' + Day2 + '/OUT/'  # OUT目录路径
        MnfstFilePath = OutDirPath + '舱单 - 副本.xlsx'  # 舱单副本表格文件路径
        import os
        if os.path.exists(MnfstFilePath) == False:  # 舱单副本表格文件不存在
            self.MsgLabel.setText("舱单副本表格文件不存在！！")
            return
        Day2DirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/航班/' + Day2 + '/'  # Day2目录路径
        CBAFilePath = Day2DirPath + 'CBA.xlsx'  # CBA表格文件路径
        if os.path.exists(CBAFilePath) == False:  # CBA表格文件不存在
            self.MsgLabel.setText("CBA表格文件不存在！！")
            return
        Month1DirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/1/'  # Month1目录路径
        from PyQt5.QtCore import QLocale
        Day2MonthEYear2 = QLocale(QLocale.English).toString(self.DateDE.date(), 'ddMMMyy').upper()  # 2位数字日 + 大写英语缩写月 + 2位数字年
        PRELOADFilePath = Month1DirPath + 'Booking list MS952 ' + Day2MonthEYear2 + ' preload.xlsx'  # BookingList表格文件路径
        if os.path.exists(PRELOADFilePath) == False:  # BookingList表格文件不存在
            self.MsgLabel.setText("BookingList表格文件不存在！！")
            return
        from ReadExcl.Mnfst.Function import ReadFltMnfstST
        ReadFltMnfstST(MnfstFilePath)  # 读取舱单副本表格航班舱单页
        from ReadExcl.CBA.Function import ReadCBAST
        ReadCBAST(CBAFilePath)  # 读取CBA表格CBA页
        from ReadExcl.CBA.Function import ReadDim
        ReadDim()  # 读取Dim
        Date = self.DateDE.date().toString('yyyy/M/d')  # 4位数字年 + 1位数字月 + 1位数字日
        from WritExcl.PRELOAD.Function import WritPRELOADST
        WritPRELOADST(PRELOADFilePath, Date)  # 写PRELOAD页
        self.MsgLabel.setText("PRELOAD录入完成")

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

    def UCM951Fctn(self):  # UCM951功能
        self.MsgLabel.setText("UCM951运行中")
        self.MsgLabel.repaint()  # MsgLabel重绘
        Date = self.DateDE.date().toString('yyMMdd')  # 2位数字年2位数字月2位数字日
        import datetime
        DateDT = datetime.datetime.strptime(Date, '%y%m%d')  # 日期字符串转日期格式
        DateDT = DateDT + datetime.timedelta(days=1)  # 日期加1天
        Date = DateDT.strftime('%y%m%d')  # 日期格式转日期字符串
        Year2Month2 = Date[:4]  # 2位数字年 + 2位数字月
        Day2 = Date[4:]  # 2位数字日
        InDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/航班/' + Day2 + '/IN/'  # IN目录路径
        UCM951CFilePath = InDirPath + 'UCM.txt'  # UCM951C文件路径
        import os
        if os.path.exists(UCM951CFilePath) == False:  # UCM951C文件不存在
            self.MsgLabel.setText("UCM951C文件不存在！！")
            return
        ULDStkDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/5/'  # ULDStock目录路径
        Day2MonthEA = DateDT.strftime('%d%b').upper()  # 日期格式转日期字符串并大写
        ULDStkFileName = Day2MonthEA + ' PVG ULD STOCK.xlsx'  # ULDStock文件名
        ULDStkFilePath = ULDStkDirPath + ULDStkFileName  # ULDStock文件路径
        if os.path.exists(ULDStkFilePath) == False:  # ULDStock文件不存在
            self.MsgLabel.setText("ULDStock文件不存在！！")
            return
        from ReadTXT.UCM951C.Function import ReadUCM
        ReadUCM(UCM951CFilePath)  # 读取UCM
        from WritExcl.ULDStk.Function import ChkUCMULDStkST
        ChkUCMULDStkST(ULDStkFilePath)  # 检查UCMULD在ULD Stock页
        self.MsgLabel.setText("UCM951比对完成")

    def ULD952Fctn(self):  # ULD952功能
        self.MsgLabel.setText("ULD952运行中")
        self.MsgLabel.repaint()  # MsgLabel重绘
        Year2Month2 = self.DateDE.date().toString('yyMM')  # 2位数字年 + 2位数字月
        Day2 = self.DateDE.date().toString('dd')  # 2位数字日
        OutDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/航班/' + Day2 + '/OUT/'  # OUT目录路径
        CPM952FilePath = OutDirPath + 'CPM.txt'  # CPM952文件路径
        import os
        if os.path.exists(CPM952FilePath) == False:  # CPM952文件不存在
            self.MsgLabel.setText("CPM952文件不存在！！")
            return
        ULDStkDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/5/'  # ULDStock目录路径
        from PyQt5.QtCore import QLocale
        Day2MonthEA = QLocale(QLocale.English).toString(self.DateDE.date(), 'ddMMM').upper()  # 2位数字日 + 大写英语缩写月
        ULDStkFileName = Day2MonthEA + ' PVG ULD STOCK.xlsx'  # ULDStock文件名
        ULDStkFilePath = ULDStkDirPath + ULDStkFileName  # ULDStock文件路径
        if os.path.exists(ULDStkFilePath) == False:  # ULDStock文件不存在
            self.MsgLabel.setText("ULDStock文件不存在！！")
            return
        from ReadTXT.CPM.Function import ReadCPM
        ReadCPM(CPM952FilePath)  # 读取CPM951,返回是否有叠板
        from ReadTXT.CPM.Function import ReadCPMULD
        ReadCPMULD()  # 读取CPMULDLst
        Year4Month1Day1 = self.DateDE.date().toString('yyyy-M-d')  # 4位数字年1位数字月1位数字日
        from WritExcl.ULDStk.Function import DelULDStkST
        DelULDStkST(ULDStkFilePath, Year4Month1Day1)  # 删除集装器在ULD Stock页
        self.MsgLabel.setText("ULD952删除完成")

    def LWS952Fctn(self):  # LWS952功能
        self.MsgLabel.setText("LWS952运行中")
        self.MsgLabel.repaint()  # MsgLabel重绘
        Year2Month2 = self.DateDE.date().toString('yyMM')  # 2位数字年 + 2位数字月
        Day2 = self.DateDE.date().toString('dd')  # 2位数字日
        OutDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/航班/' + Day2 + '/OUT/'  # OUT目录路径
        from PyQt5.QtCore import QLocale
        Day2MonthEA = QLocale(QLocale.English).toString(self.DateDE.date(), 'ddMMM').upper()  # 2位数字日 + 大写英语缩写月
        # LWS952FileName = 'DLWS ' + Day2MonthEA + '.pdf'  # LWS952文件名
        LWS952FileName = 'MS952 ' + Day2MonthEA + ' DLWS.pdf'  # LWS952文件名
        LWS952FilePath = OutDirPath + LWS952FileName  # LWS952文件路径
        import os
        if os.path.exists(LWS952FilePath) == False:  # LWS952文件不存在
            self.MsgLabel.setText("LWS952文件不存在！！")
            return
        ULDStkDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/5/'  # ULDStock目录路径
        ULDStkFileName = Day2MonthEA + ' PVG ULD STOCK.xlsx'  # ULDStock文件名
        ULDStkFilePath = ULDStkDirPath + ULDStkFileName  # ULDStock文件路径
        if os.path.exists(ULDStkFilePath) == False:  # ULDStock文件不存在
            self.MsgLabel.setText("ULDStock文件不存在！！")
            return
        # UnilodeSCMFilePath = ULDStkDirPath + 'Unilode SCM.txt'  # Unilode SCM文件路径
        # if os.path.exists(UnilodeSCMFilePath) == False:  # Unilode SCM文件不存在
        #     self.MsgLabel.setText("Unilode SCM文件不存在！！")
        #     return
        from ReadPDF.LWS.Function import ReadLWS
        ReadLWS(LWS952FilePath)  # 读取LWS
        Year4Month1Day1 = self.DateDE.date().toString('yyyy-M-d')  # 4位数字年1位数字月1位数字日
        from WritExcl.ULDStk.Function import DelULDStkST
        DelULDStkST(ULDStkFilePath, Year4Month1Day1)  # 删除集装器在ULD Stock页
        # from ReadExcl.ULDStk.Function import ReadUnilodeST
        # UnilodeSCM = ReadUnilodeST(ULDStkFilePath, Day2MonthEA)  # 读取Unilode页,返回Unilode SCM
        # from WritTXT.SCM.Function import WritSCM
        # WritSCM(UnilodeSCMFilePath, UnilodeSCM)  # 写Unilode SCM文件
        self.MsgLabel.setText("LWS952更新完成")

    def AKE952Fctn(self):  # AKE952功能
        self.MsgLabel.setText("AKE952运行中")
        self.MsgLabel.repaint()  # MsgLabel重绘
        Year2Month2 = self.DateDE.date().toString('yyMM')  # 2位数字年 + 2位数字月
        Day2 = self.DateDE.date().toString('dd')  # 2位数字日
        OutDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/航班/' + Day2 + '/OUT/'  # OUT目录路径
        AKE952FilePath = OutDirPath + 'AKE.txt'  # AKE952文件路径
        import os
        if os.path.exists(AKE952FilePath) == False:  # AKE952文件不存在
            self.MsgLabel.setText("AKE952文件不存在！！")
            return
        ULDStkDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/5/'  # ULDStock目录路径
        from PyQt5.QtCore import QLocale
        Day2MonthEA = QLocale(QLocale.English).toString(self.DateDE.date(), 'ddMMM').upper()  # 2位数字日 + 大写英语缩写月
        ULDStkFileName = Day2MonthEA + ' PVG ULD STOCK.xlsx'  # ULDStock文件名
        ULDStkFilePath = ULDStkDirPath + ULDStkFileName  # ULDStock文件路径
        if os.path.exists(ULDStkFilePath) == False:  # ULDStock文件不存在
            self.MsgLabel.setText("ULDStock文件不存在！！")
            return
        from ReadTXT.AKE952.Function import ReadAKE952
        ReadAKE952(AKE952FilePath)  # 读取AKE952
        from ReadTXT.CPM.Function import ReadCPMULD
        ReadCPMULD()  # 读取CPMULDLst
        Year4Month1Day1 = self.DateDE.date().toString('yyyy-M-d')  # 4位数字年1位数字月1位数字日
        from WritExcl.ULDStk.Function import DelULDStkST
        DelULDStkST(ULDStkFilePath, Year4Month1Day1)  # 删除集装器在ULD Stock页
        self.MsgLabel.setText("AKE952删除完成")

    def RentalMnfstFctn(self):  # 租板舱单功能
        self.MsgLabel.setText("租板舱单运行中")
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
        from WritExcl.RentalMnfst.Function import ReadMnfstLst
        ReadMnfstLst()  # 读取舱单对象列表
        from WritExcl.RentalMnfst.Function import WritRentalMnfstST
        WritRentalMnfstST(MnfstFilePath)  # 写租板舱单页
        self.MsgLabel.setText("租板舱单录入完成")

    def SCM1Fctn(self):  # SCM1功能
        self.MsgLabel.setText("SCM1运行中")
        self.MsgLabel.repaint()  # MsgLabel重绘
        SCMPACTLFilePath = 'C:/Files/MS/ULD/SCM/SCM PACTL.TXT'  # SCM PACTL文件路径
        import os
        if os.path.exists(SCMPACTLFilePath) == False:  # SCM PACTL文件不存在
            self.MsgLabel.setText("SCM PACTL文件不存在！！")
            return
        Year2Month2 = self.DateDE.date().toString('yyMM')  # 2位数字年 + 2位数字月
        ULDStkDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/5/'  # ULDStock目录路径
        from PyQt5.QtCore import QLocale
        Day2MonthEA = QLocale(QLocale.English).toString(self.DateDE.date(), 'ddMMM').upper()  # 2位数字日 + 大写英语缩写月
        ULDStkFileName = Day2MonthEA + ' PVG ULD STOCK.xlsx'  # ULDStock文件名
        ULDStkFilePath = ULDStkDirPath + ULDStkFileName  # ULDStock文件路径
        if os.path.exists(ULDStkFilePath) == False:  # ULDStock文件不存在
            self.MsgLabel.setText("ULDStock文件不存在！！")
            return
        from ReadTXT.SCM.Function import ReadSCM
        ReadSCM(SCMPACTLFilePath)  # 读取SCM
        from WritExcl.ULDStk.Function import ChkSCMULDStkST
        ChkSCMULDStkST(ULDStkFilePath)  # 检查SCM在ULD Stock页
        self.MsgLabel.setText("SCM1完成")

    def SCM2Fctn(self):  # SCM2功能
        self.MsgLabel.setText("SCM2运行中")
        self.MsgLabel.repaint()  # MsgLabel重绘
        SCMPACTLFilePath = 'C:/Files/MS/ULD/SCM/SCM PACTL.TXT'  # SCM PACTL文件路径
        import os
        if os.path.exists(SCMPACTLFilePath) == False:  # SCM PACTL文件不存在
            self.MsgLabel.setText("SCM PACTL文件不存在！！")
            return
        Year2Month2 = self.DateDE.date().toString('yyMM')  # 2位数字年 + 2位数字月
        ULDStkDirPath = 'C:/Files/MS/日常/' + Year2Month2 + '/5/'  # ULDStock目录路径
        from PyQt5.QtCore import QLocale
        Day2MonthEA = QLocale(QLocale.English).toString(self.DateDE.date(), 'ddMMM').upper()  # 2位数字日 + 大写英语缩写月
        ULDStkFileName = Day2MonthEA + ' PVG ULD STOCK.xlsx'  # ULDStock文件名
        ULDStkFilePath = ULDStkDirPath + ULDStkFileName  # ULDStock文件路径
        import os
        if os.path.exists(ULDStkFilePath) == False:  # ULDStock文件不存在
            self.MsgLabel.setText("ULDStock文件不存在！！")
            return
        from ReadExcl.ULDStk.Function import ReadULDStkST
        SCM = ReadULDStkST(ULDStkFilePath)  # 读取ULD Stock页,返回SCM
        SCMFilePath = 'C:/Files/MS/ULD/SCM/SCM.TXT'  # SCM文件路径
        from WritTXT.SCM.Function import WritSCM
        WritSCM(SCMFilePath, SCM)  # 写SCM文件
        self.MsgLabel.setText("SCM2完成")