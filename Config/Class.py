class Cfg():  # 配置文件类
    PMCTup2 = ()  # PMC起始结束行元组
    PAGTup2 = ()  # PAG起始结束行元组
    PLATup2 = ()  # PLA起始结束行元组
    AKETup2 = ()  # AKE起始结束行元组

    def __init__(self):  # 构造函数
        from configobj import ConfigObj
        Cfg = ConfigObj('C:/Files/Python/MS/Config/ULDStk.ini', encoding='UTF-8')
        self.PMCTup2 = (int(Cfg['PMC']['r1']), int(Cfg['PMC']['r2']))  # PMC起始结束行
        self.PAGTup2 = (int(Cfg['PAG']['r1']), int(Cfg['PAG']['r2']))  # PAG起始结束行
        self.PLATup2 = (int(Cfg['PLA']['r1']), int(Cfg['PLA']['r2']))  # PLA起始结束行
        self.AKETup2 = (int(Cfg['AKE']['r1']), int(Cfg['AKE']['r2']))  # AKE起始结束行