class Shpmt:  # Shpmt类
    AWBNo = ''  # 运单号
    SHC = ''  # 特殊操作代码
    ManDesc = ''  # 品名
    Pcs = 0  # 件数
    LeftPcs = 0  # 剩余件数
    Weight = 0  # 重量
    LeftWt = 0  # 剩余重量
    ChgWt = 0  # 计费重量
    LeftChgWt = 0  # 剩余计费重量
    Vol = 0  # 体积
    LeftVol = 0  # 剩余体积

    def __init__(self, AWBNo, SHC, ManDesc, Pcs, Weight, ChgWt, Vol):  # 构造函数
        self.AWBNo = AWBNo  # 得到运单号
        self.SHC = SHC  # 得到特殊操作代码
        self.ManDesc = ManDesc  # 得到品名
        self.Pcs = int(Pcs)  # 得到件数
        self.LeftPcs = self.Pcs  # 得到剩余件数
        self.Weight = float(Weight)  # 得到重量
        self.LeftWt = self.Weight  # 得到剩余重量
        self.ChgWt = float(ChgWt)  # 得到计费重量
        self.LeftChgWt = self.ChgWt  # 得到剩余计费重量
        self.Vol = float(Vol)  # 得到体积
        self.LeftVol = self.Vol  # 得到剩余体积

class ShpmtULD:  # ShpmtULD类
    Type = ''  # 类型
    No = ''  # 号码
    Pcs = 0  # 件数
    Weight = 0  # 重量
    Vol = 0  # 体积
    ChgWt = 0  # 计费重量



