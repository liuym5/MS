def ReadLWS(Path):  # 读取LWS
    from ReadPDF.Function import ReadPage1PDF
    LWS = ReadPage1PDF(Path)  # 返回LWS952第1页PDF文件文本
    LWS = LWS.split('-')  # 以-分割字符串
    TypeTup = ('AKE', 'PMC', 'PLA', 'PAG', 'P1P')  # 类型元组
    i = 0  # LWS字符串列表下标
    for item in LWS:  # 遍历LWS字符串列表
        from ReadTXT.CPM.Function import FindType
        Type, j = FindType(TypeTup, item)  # 返回类型,类型首字母下标
        if j > -1:  # 找到类型
            No = LWS[i + 1]  # 号
            Owner = LWS[i + 2][:2]  # 所有人
            from ReadTXT.UCM951.Class import UCMULD
            UCMULDTmp = UCMULD(Type, No, Owner)  # 创建UCMULD对象
            from ReadTXT.UCM951.Variable import UCMULDLst
            UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表
        i += 1  # LWS字符串列表下标加1