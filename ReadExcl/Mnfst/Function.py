def ReadFltMnfstST(MnfstPath):  # 读取舱单副本表格航班舱单页
    import pandas as pd
    df = pd.read_excel(MnfstPath, header=None)  # 读取舱单副本表格航班舱单页
    for r in range(len(df)):  # 遍历所有行
        AWBNo = df.iloc[r][1]  # 得到运单号
        SHC = df.iloc[r][5]  # 得到特殊操作代码
        ManDesc = df.iloc[r][6]  # 得到品名
        Pcs = df.iloc[r][7]  # 得到件数
        Weight = df.iloc[r][8]  # 得到重量
        ChgWt = df.iloc[r][9]  # 得到计费重量
        Vol = df.iloc[r][10]  # 得到体积
        from ReadExcl.Mnfst.Class import Shpmt
        Row = Shpmt(AWBNo, SHC, ManDesc, Pcs, Weight, ChgWt, Vol)  # 创建当前行Shpmt对象
        from ReadExcl.Mnfst.Variable import MnfstLst
        MnfstLst.append(Row)  # 添加到航班舱单Shpmt对象列表
    # for i in range(len(MnfstLst)):  # 测试代码,遍历列表
    #     print(MnfstLst[i].__dict__)  # 打印列表

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
                j = r + 3  # 得到运单号行号
                C1 = str(df.iloc[j][1])  # 得到第1列字符串
                while C1 != 'Total':  # 不是Total字符串
                    if C1.find('077-') > -1:  # 找到077-字符串
                        AWBNo = C1  # 得到运单号
                    j += 1  # 行号加1
                    C1 = str(df.iloc[j][1])  # 得到第1列字符串
                break
    pd.set_option('display.max_rows', None)  # 显示所有行
    pd.set_option('expand_frame_repr', False)  # 显示所有列
    print(df)  # 打印数据框架
