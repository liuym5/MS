def ReadFltMnfstST(MnfstPath):  # 读取舱单副本表格航班舱单页
    import pandas as pd
    df = pd.read_excel(MnfstPath, header=None)  # 读取舱单副本表格航班舱单页
    for r in range(len(df)):  # 遍历所有行
        AWBNo = df.iloc[r][1]  # 得到运单号
        Dest = df.iloc[r][3]  # 得到目的地
        SHC = df.iloc[r][5]  # 得到特殊操作代码
        ManDesc = df.iloc[r][6]  # 得到品名
        Pcs = df.iloc[r][7]  # 得到件数
        Weight = df.iloc[r][8]  # 得到重量
        ChgWt = df.iloc[r][9]  # 得到计费重量
        Vol = df.iloc[r][10]  # 得到体积
        from ReadExcl.Mnfst.Class import Shpmt
        ShpmtTmp = Shpmt(AWBNo, Dest, SHC, ManDesc, Pcs, Weight, ChgWt, Vol)  # 创建Shpmt对象
        from ReadExcl.Mnfst.Variable import MnfstLst
        MnfstLst.append(ShpmtTmp)  # 添加到航班舱单Shpmt对象列表

def ReadULDMnfstST(MnfstPath):  # 读取舱单副本表格ULD舱单页
    TypeTup = ('PMC', 'PAG', 'PLA', 'AKE', 'P1P')  # 类型元组
    import pandas as pd
    df = pd.read_excel(MnfstPath, sheet_name=1, header=None)  # 读取舱单副本表格ULD舱单页
    for r in range(len(df)):  # 遍历所有行
        C0 = str(df.iloc[r][0])  # 得到第0列字符串
        for i in range(len(TypeTup)):  # 遍历类型元组
            TypeTmp = TypeTup[i]  # 得到类型临时
            j = C0.find(TypeTmp)  # 找类型
            if j > -1:  # 找到类型
                Type = TypeTmp  # 得到类型
                No = C0[j + 4 : j + 9]  # 得到号码
                Owner = C0[j + 10 : j + 12]  # 得到所有人
                j = r + 3  # 得到运单号行号
                C1 = str(df.iloc[j][1])  # 得到第1列字符串
                while C1 != 'Total':  # 不是Total字符串
                    if C1.find('077-') > -1:  # 找到077-字符串
                        AWBNo = C1  # 得到运单号
                        Pcs = int(df.iloc[j][2].split('/')[0])  # 得到件数
                        Weight = float(df.iloc[j][4])  # 得到重量
                        from ReadExcl.Mnfst.Class import ShpmtULD
                        ShpmtULDTmp = ShpmtULD(Type, No, Owner, Pcs, Weight)  # 创建ShpmtULD对象
                        from ReadExcl.Mnfst.Variable import MnfstLst
                        ShpmtTmp = FindAWBNo(MnfstLst, AWBNo)  # 返回该运单号货物对象
                        ShpmtTmp.AddULD(ShpmtULDTmp)  # 添加集装器
                    j += 1  # 行号加1
                    C1 = str(df.iloc[j][1])  # 得到第1列字符串
                break

def FindAWBNo(MnfstLst, AWBNo):  # 返回该运单号货物对象
    for Shpmt in MnfstLst:  # 遍历舱单对象列表
        if Shpmt.AWBNo == AWBNo:  # 找到舱单对象相对应的运单号
            return Shpmt  # 返回货物对象