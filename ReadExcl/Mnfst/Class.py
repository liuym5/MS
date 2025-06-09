class Shpmt:  # Shpmt类
    Seq = ''  # 序号
    AWBNo = ''  # 运单号
    Dest = ''  # 目的地
    SHC = ''  # 特殊操作代码
    ManDesc = ''  # 品名
    Pcs = 0  # 件数
    Weight = 0  # 重量
    ChgWt = 0  # 计费重量
    LeftChgWt = 0  # 剩余计费重量
    Vol = 0  # 体积
    LeftVol = 0  # 剩余体积
    Dim = ''  # 尺寸
    ULDLst = []  # 集装器对象列表

    def __init__(self, Seq, AWBNo, Dest, SHC, ManDesc, Pcs, Weight, ChgWt, Vol):  # 构造函数
        self.Seq = Seq  # 序号
        self.AWBNo = AWBNo  # 运单号
        self.Dest = Dest  # 目的地
        self.SHC = SHC  # 特殊操作代码
        self.ManDesc = ManDesc  # 品名
        self.Pcs = int(Pcs)  # 件数
        self.Weight = float(Weight)  # 重量
        self.LeftWt = self.Weight  # 剩余重量
        self.ChgWt = float(ChgWt)  # 计费重量
        self.LeftChgWt = self.ChgWt  # 剩余计费重量
        self.Vol = float(Vol)  # 体积
        self.LeftVol = self.Vol  # 剩余体积
        self.ULDLst = []  # 创建集装器对象列表新内存地址并引用

    def AddULD(self, ShptULD):  # 添加集装器
        self.LeftWt, \
        self.LeftVol, \
        self.LeftChgWt = ShptULD.GetVolChgWt(self.LeftWt, self.LeftVol, self.LeftChgWt)  # 返回剩余重量,剩余体积,剩余计费重量
        self.ULDLst.append(ShptULD)  # 添加货物集装器对象到集装器对象列表

class ShpmtULD:  # ShpmtULD类
    Type = ''  # 类型
    No = ''  # 号
    Owner = ''  # 所有人
    ULDNo = ''  # 集装器号
    FullULDNo = ''  # 集装器号全称
    Pcs = 0  # 件数
    Weight = 0  # 重量
    Vol = 0  # 体积
    ChgWt = 0  # 计费重量

    def __init__(self, Type, No, Owner, Pcs, Weight):  # 构造函数
        self.Type = Type  # 类型
        self.No = No  # 号
        self.Owner = Owner  # 所有人
        self.ULDNo = self.Type + self.No  # 集装器号
        self.FullULDNo = self.Type + self.No + self.Owner # 得到集装器号全称
        self.Pcs = int(Pcs)  # 件数
        self.Weight = float(Weight)  # 重量

    def GetVolChgWt(self, LeftWt, LeftVol, LeftChgWt):  # 返回剩余重量,剩余体积,剩余计费重量
        if self.Weight < LeftWt:  # 还有剩余重量
            self.Vol = round(self.Weight / LeftWt * LeftVol, 3)  # 体积
            self.ChgWt = round(self.Weight / LeftWt * LeftChgWt, 2)  # 计费重量
            LeftWt = round(LeftWt - self.Weight, 2)  # 剩余重量
            LeftVol = round(LeftVol - self.Vol, 3)  # 剩余体积
            LeftChgWt = round(LeftChgWt - self.ChgWt, 2)  # 剩余计费重量
            return LeftWt, LeftVol, LeftChgWt  # 返回剩余重量,剩余体积,剩余计费重量
        self.Vol = LeftVol  # 体积
        self.ChgWt = LeftChgWt  # 计费重量
        LeftWt = LeftWt - self.Weight  # 剩余重量
        LeftVol = LeftVol - self.Vol  # 剩余体积
        LeftChgWt = LeftChgWt - self.ChgWt  # 剩余计费重量
        return LeftWt, LeftVol, LeftChgWt  # 返回剩余重量,剩余体积,剩余计费重量