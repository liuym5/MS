def ReadUCM(Path):  # 读取UCM
    from ReadTXT.Function import ReadTXT
    UCM = ReadTXT(Path)  # 返回UCM TXT文件文本
    UCM = UCM.split('.')  # 以.分割字符串
    TypeTup = ('AKE', 'PMC', 'PLA', 'PAG', 'P1P')  # 类型元组
    for item in UCM:  # 遍历UCM字符串列表
        from ReadTXT.CPM.Function import FindType
        Type, i = FindType(TypeTup, item)  # 返回类型,类型首字母下标
        if i > -1:  # 找到类型
            No = item[3:8]  # 号
            Owner = item[8:10]  # 所有人
            from ReadTXT.UCM951 import UCMULD
            UCMULDTmp = UCMULD(Type, No, Owner)  # 创建UCMULD对象
            from ReadTXT.UCM951.Variable import UCMULDLst
            UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表