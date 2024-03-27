def ReadUCM(UCM):  # 读取UCM
    UCM = UCM.replace('\n', '').replace('SI', '')  # 去除换行符和SI
    UCM = UCM.split('.') # 以.分割字符串
    TypeTup = ('AKE', 'PMC', 'PLA', 'PAG', 'P1P')  # 类型元组
    No = ''  # 号
    Owner = ''  # 所有人
    for Pos in UCM:  # 遍历UCM字符串
        from ReadPDF.CPM.Function import FindType
        Type, i = FindType(TypeTup, Pos)  # 返回类型,类型首字母下标
        if i > -1:  # 找到类型
            if Pos[-2:] in ('MS', 'DS'):  # 所有人为MS或DS
                Owner = Pos[-2:]  # 所有人为MS或DS
                Pos = Pos.split(Owner)  # 以所有人分割字符串
                No = Pos[3:]  # 号
            else:  # 无所有人
                No = Pos[3:]  # 号
            if len(No) == 5:  # 号为5位
                from ReadPDF.UCM.Class import UCMULD
                UCMULDTmp = UCMULD(Type, No, Owner)  # 创建UCMULD对象
                from ReadPDF.UCM.Variable import UCMULDLst
                UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表
            elif len(No) == 4:  # 号为4位
                No = '0' + No  # 号前面加个0补充到5位
                from ReadPDF.UCM.Class import UCMULD
                UCMULDTmp = UCMULD(Type, No, Owner)  # 创建UCMULD对象
                from ReadPDF.UCM.Variable import UCMULDLst
                UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表
    DelULD()  # 删除重复ULD

def DelULD():  # 删除重复ULD
    from ReadPDF.CPM.Variable import CPMULDLst
    for CPMULD in CPMULDLst:  # 遍历CPMULDLst
        i = 0  # UCMULDLst下标
        from ReadPDF.UCM.Variable import UCMULDLst
        for UCMULD in UCMULDLst:  # 遍历UCMULDLst
               if CPMULD.No == UCMULD.No:  # No相同
                   del UCMULDLst[i]  # 删除相同No列表项
                   break
               i += 1  # 下标加1