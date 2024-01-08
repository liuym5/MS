class Shpmt:  # Shpmt类
    AWBNo = ''  # 运单号
    SHC = ''  # 特殊操作代码
    ManDesc = ''  # 品名
    Pcs = 0  # 件数
    Weight = 0  # 重量
    ChgWt = 0  # 计费重量
    LeftChgWt = 0  # 剩余计费重量
    Vol = 0  # 体积
    LeftVol = 0  # 剩余体积
    ULDLst = []  # 集装器对象列表

    def __init__(self, AWBNo, SHC, ManDesc, Pcs, Weight, ChgWt, Vol):  # 构造函数
        self.AWBNo = AWBNo  # 得到运单号
        self.SHC = SHC  # 得到特殊操作代码
        self.ManDesc = ManDesc  # 得到品名
        self.Pcs = int(Pcs)  # 得到件数
        self.Weight = float(Weight)  # 得到重量
        self.LeftWt = self.Weight  # 得到剩余重量
        self.ChgWt = float(ChgWt)  # 得到计费重量
        self.LeftChgWt = self.ChgWt  # 得到剩余计费重量
        self.Vol = float(Vol)  # 得到体积
        self.LeftVol = self.Vol  # 得到剩余体积
        self.ULDLst = []  # 创建集装器对象列表新内存地址并引用

    def AddULD(self, ShptULD):  # 添加集装器
        self.LeftWt, \
        self.LeftVol, \
        self.LeftChgWt = ShptULD.GetVolChgWt(self.LeftWt, self.LeftVol, self.LeftChgWt)  # 返回剩余重量,剩余体积,剩余计费重量
        self.ULDLst.append(ShptULD)  # 添加货物集装器对象到集装器对象列表

class ShpmtULD:  # ShpmtULD类
    Type = ''  # 类型
    No = ''  # 号
    ULDNo = ''  # 集装器号
    Pcs = 0  # 件数
    Weight = 0  # 重量
    Vol = 0  # 体积
    ChgWt = 0  # 计费重量

    def __init__(self, Type, No, Pcs, Weight):  # 构造函数
        self.Type = Type  # 得到类型
        self.No = No  # 得到号
        self.ULDNo = self.Type + self.No  # 得到集装器号
        self.Pcs = int(Pcs)  # 得到件数
        self.Weight = float(Weight)  # 得到重量

    def GetVolChgWt(self, LeftWt, LeftVol, LeftChgWt):  # 返回剩余重量,剩余体积,剩余计费重量
        if self.Weight < LeftWt:  # 还有剩余重量
            self.Vol = round(self.Weight / LeftWt * LeftVol, 2)  # 得到体积
            self.ChgWt = round(self.Weight / LeftWt * LeftChgWt, 2)  # 得到计费重量
            LeftWt = round(LeftWt - self.Weight, 2)  # 得到剩余重量
            LeftVol = round(LeftVol - self.Vol, 2)  # 得到剩余体积
            LeftChgWt = round(LeftChgWt - self.ChgWt, 2)  # 得到剩余计费重量
            return LeftWt, LeftVol, LeftChgWt  # 返回剩余重量,剩余体积,剩余计费重量
        self.Vol = LeftVol  # 得到体积
        self.ChgWt = LeftChgWt  # 得到计费重量
        LeftWt = LeftWt - self.Weight  # 得到剩余重量
        LeftVol = LeftVol - self.Vol  # 得到剩余体积
        LeftChgWt = LeftChgWt - self.ChgWt  # 得到剩余计费重量
        return LeftWt, LeftVol, LeftChgWt  # 返回剩余重量,剩余体积,剩余计费重量