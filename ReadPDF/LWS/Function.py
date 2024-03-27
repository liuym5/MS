def ReadLWS(LWS):  # 读取LWS
    LWS = LWS.split('-')  # 以-分割字符串
    TypeTup = ('AKE', 'PMC', 'PLA', 'PAG', 'P1P')  # 类型元组
    i = 0  # LWS字符串列表下标
    No = ''  # 号
    Owner = ''  # 所有人
    for Pos in LWS:  # 遍历LWS字符串列表
        from ReadPDF.CPM.Function import FindType
        Type, j = FindType(TypeTup, Pos)  # 返回类型,类型首字母下标
        if j > -1:  # 找到类型
            No = LWS[i + 1]  # 号
            Owner = LWS[i + 2][:2]  # 所有人
            from ReadPDF.UCM.Class import UCMULD
            UCMULDTmp = UCMULD(Type, No, Owner)  # 创建UCMULD对象
            from ReadPDF.UCM.Variable import UCMULDLst
            UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表
        i += 1  # LWS字符串列表下标加1