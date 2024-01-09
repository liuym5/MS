from ReadExcl.Mnfst.Class import ShpmtULD, Shpmt

class DSULD(ShpmtULD):  # DSULD类,继承ShpmtULD类
    Ser = 0  # 序号
    FwdTo = ''  # 运往
    ShptLst = []  # 货物对象列表

    def __init__(self, Ser, Type, No, Owner, FwdTo, DSShpt):  # 构造函数
        self.Ser = Ser  # 得到序号
        self.Type = Type  # 得到类型
        self.No = No  # 得到号
        self.Owner = Owner  # 得到所有人
        self.FwdTo = FwdTo  # 得到运往
        self.ShptLst = []  # 创建货物对象列表新内存地址并引用
        self.ShptLst.append(DSShpt)  # 添加DS货物对象到货物对象列表

    def AddShpt(self, DSShpt):  # 添加货物
        self.ShptLst.append(DSShpt)  # 添加DS货物对象到货物对象列表

class DSShpmt(Shpmt):  # DSShpmt类,继承Shpmt类
    TtlPcs = 0  # 总件数

    def __init__(self, AWBNo, Dest, Pcs, TtlPcs, Weight):  # 构造函数
        self.AWBNo = AWBNo  # 得到运单号
        self.Dest = Dest  # 得到目的地
        self.Pcs = Pcs  # 得到件数
        self.TtlPcs = TtlPcs  # 得到总件数
        self.Weight = Weight  # 得到重量