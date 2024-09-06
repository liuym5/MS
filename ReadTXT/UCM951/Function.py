def ReadUCM(Path):  # 读取UCM
    from ReadTXT.Function import ReadTXT
    UCM = ReadTXT(Path)  # 返回UCM951文件文本
    UCM = UCM.replace('\n', '').replace('SI', '').rstrip()  # 去除换行符和SI和末尾空格
    UCM = UCM.split('.') # 以.分割字符串
    TypeTup = ('AKE', 'PMC', 'PAG', 'PLA', 'PAJ', 'P1P')  # 类型元组
    for item in UCM:  # 遍历UCM字符串列表
        from ReadTXT.CPM.Function import FindType
        Type, i = FindType(TypeTup, item)  # 返回类型,类型首字母下标
        if i > -1:  # 找到类型
            if item[-2:] in ('R7', 'R9', 'C6', 'K8', 'KE'):  # 所有人为R7,R9,C6,K8,KE
                Owner = item[-2:]  # 所有人
                item = item.split(Owner)  # 以所有人分割字符串
                No = item[0][3:]  # 号
            else:  # 无所有人
                No = item[3:]  # 号
                Owner = ''  # 所有人
            if len(No) == 5:  # 号为5位
                from ReadTXT.UCM951.Class import UCMULD
                UCMULDTmp = UCMULD(Type, No, Owner)  # 创建UCMULD对象
                from ReadTXT.UCM951.Variable import UCMULDLst
                UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表
            elif len(No) == 4:  # 号为4位
                No = '0' + No  # 号前面加个0补充到5位
                from ReadTXT.UCM951.Class import UCMULD
                UCMULDTmp = UCMULD(Type, No, Owner)  # 创建UCMULD对象
                from ReadTXT.UCM951.Variable import UCMULDLst
                UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表
    import operator
    cmpfun = operator.attrgetter('No')  # 参数为排序依据的属性，可以有多个，这里优先id，使用时按需求改换参数即可
    from ReadTXT.UCM951.Variable import UCMULDLst
    UCMULDLst.sort(key=cmpfun)  # 使用时改变列表名即可
    DelULD()  # 删除重复ULD

def DelULD():  # 删除重复ULD
    from ReadTXT.CPM.Variable import CPMULDLst
    for cpmuld in CPMULDLst:  # 遍历CPMULDLst
        i = 0  # UCMULDLst下标
        from ReadTXT.UCM951.Variable import UCMULDLst
        for ucmuld in UCMULDLst:  # 遍历UCMULDLst
               if cpmuld.No == ucmuld.No:  # No相同
                   del UCMULDLst[i]  # 删除相同No列表项
                   break
               i += 1  # 下标加1