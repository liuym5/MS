from ReadExcl.Mnfst.Class import Shpmt

class CBADim(Shpmt):  # CBADim类,继承Shpmt类
    Dim = ''  # 尺寸

    def __init__(self, AWBNo, Dim):  # 构造函数
        self.AWBNo = AWBNo  # 运单号
        self.Dim = Dim  # 尺寸