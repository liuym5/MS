from ReadExcl.Mnfst.Class import ShpmtULD

class UCMULD(ShpmtULD):  # UCMULD类,继承ShpmtULD类
    def __init__(self, Type, No, Owner):  # 构造函数
        self.Type = Type  # 类型
        self.No = No  # 号
        self.Owner = Owner  # 所有人
        self.ULDNo = self.Type + self.No  # 集装器号
        self.FullULDNo = self.Type + self.No + self.Owner  # 集装器号全称