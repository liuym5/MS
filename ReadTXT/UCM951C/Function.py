def ReadUCM(Path):  # 读取UCM
    from ReadTXT.Function import ReadTXT
    UCM = ReadTXT(Path)  # 返回UCM TXT文件文本
    UCM = UCM.split('.')  # 以.分割字符串
    TypeTup = ('AKE', 'PMC', 'PLA', 'PAG', 'PAJ', 'P1P')  # 类型元组
    for item in UCM:  # 遍历UCM字符串列表
        from ReadTXT.CPM.Function import FindType
        Type, i = FindType(TypeTup, item)  # 返回类型,类型首字母下标
        if i > -1:  # 找到类型
            No = item[3:7]  # 4位号
            Owner = item[7:9]  # 所有人
            if No.isdigit() and Owner in ('MS', 'R7', 'R9', 'C6', 'K8', 'DS'):  # 4位号为数字并且所有人为MS,R7,R9,C6,K8,DS
                No = '0' + No  # 号前面加个0补充到5位
                from ReadTXT.UCM951.Class import UCMULD
                UCMULDTmp = UCMULD(Type, No, Owner)  # 创建UCMULD对象
                from ReadTXT.UCM951.Variable import UCMULDLst
                UCMULDLst.append(UCMULDTmp)  # 添加CPMULD对象到CPMULD对象列表
            No = item[3:8]  # 5位号
            Owner = item[8:10]  # 所有人
            if No.isdigit() and Owner in ('MS', 'R7', 'R9', 'C6', 'K8', 'DS'):  # 5位号为数字并且所有人为MS,R7,R9,C6,K8,DS
                from ReadTXT.UCM951.Class import UCMULD
                UCMULDTmp = UCMULD(Type, No, Owner)  # 创建UCMULD对象
                from ReadTXT.UCM951.Variable import UCMULDLst
                UCMULDLst.append(UCMULDTmp)  # 添加CPMULD对象到CPMULD对象列表