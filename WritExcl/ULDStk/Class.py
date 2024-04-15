from ReadExcl.Mnfst.Class import ShpmtULD

class ColorULD(ShpmtULD):  # ColorULD类,继承ShpmtULD类
    Font = 1  # 单元格文字颜色为1黑色
    Interior = 2  # 单元格背景颜色为2白色

    def __init__(self, No, Font, Interior):  # 构造函数
        self.No = No  # 号
        self.Font = Font  # 单元格文字颜色号
        self.Interior = Interior  # 单元格背景颜色号