def ReadSCM(Path):  # 读取SCM
    from ReadTXT.Function import ReadTXT
    SCM = ReadTXT(Path)  # 返回SCM文件文本
    SCM = SCM.replace('\n', '')  # 去除换行符
    SCM = SCM.split('/')  # 以/分割字符串
    TypeTup = ('AKE', 'PAG', 'PLA', 'PMC', 'P1P')  # 类型元组
    for item in SCM:  # 遍历SCM字符串列表
        from ReadTXT.CPM.Function import FindType
        Type, i = FindType(TypeTup, item)  # 返回类型,类型首字母下标
        if i > -1:  # 找到类型
            if Type == 'AKE':  # 类型为AKE
                FindAKE(item, i)  # 找到AKE
                CurType = 'AKE'
            elif Type == 'PAG':  # 类型为PAG
                FindPLT(item, CurType, Type)  # 找到板
                CurType = 'PAG'
            elif Type == 'PLA':  # 类型为PAG
                FindPLT(item, CurType, Type)  # 找到板
                CurType = 'PLA'
            elif Type == 'PMC':  # 类型为PAG
                FindPLT(item, CurType, Type)  # 找到板
                CurType = 'PMC'
            elif Type == 'P1P':  # 类型为PAG
                FindPLT(item, CurType, Type)  # 找到板
                CurType = 'P1P'
        else:  # 没有找到类型
            No = item[:5]  # 5位号
            if No.isdigit():  # 是5位号
                Owner = item[5:7]  # 所有人
                from ReadTXT.UCM951.Class import UCMULD
                UCMULDTmp = UCMULD(CurType, No, Owner)  # 创建UCMULD对象
                from ReadTXT.UCM951.Variable import UCMULDLst
                UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表
            else:  # 4位号
                No = '0' + item[:4]  # 4位号
                if No.isdigit():  # 是4位号
                    Owner = item[4:6]  # 所有人
                    from ReadTXT.UCM951.Class import UCMULD
                    UCMULDTmp = UCMULD(CurType, No, Owner)  # 创建UCMULD对象
                    from ReadTXT.UCM951.Variable import UCMULDLst
                    UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表

def FindAKE(item, i):  # 找到AKE
    No = item[i + 4:i + 9]  # 5位号
    if No.isdigit():  # 是5位号
        Owner = item[i + 9:i + 11]  # 所有人
        from ReadTXT.UCM951.Class import UCMULD
        UCMULDTmp = UCMULD('AKE', No, Owner)  # 创建UCMULD对象
        from ReadTXT.UCM951.Variable import UCMULDLst
        UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表
    else:  # 4位号
        No = '0' + item[i + 4:i + 8]  # 4位号
        Owner = item[i + 8:i + 10]  # 所有人
        from ReadTXT.UCM951.Class import UCMULD
        UCMULDTmp = UCMULD('AKE', No, Owner)  # 创建UCMULD对象
        from ReadTXT.UCM951.Variable import UCMULDLst
        UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表

def FindPLT(item, CurType, Type):  # 找到板
    No = item[:5]  # 5位号
    if No.isdigit():  # 是5位号
        Owner = item[5:7]  # 所有人
        from ReadTXT.UCM951.Class import UCMULD
        UCMULDTmp = UCMULD(CurType, No, Owner)  # 创建UCMULD对象
        from ReadTXT.UCM951.Variable import UCMULDLst
        UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表
    else:  # 4位号
        No = '0' + item[:4]  # 4位号
        Owner = item[4:6]  # 所有人
        from ReadTXT.UCM951.Class import UCMULD
        UCMULDTmp = UCMULD(CurType, No, Owner)  # 创建UCMULD对象
        from ReadTXT.UCM951.Variable import UCMULDLst
        UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表
    No = item[16:21]  # 5位号
    if No.isdigit():  # 是5位号
        Owner = item[21:23]  # 所有人
        from ReadTXT.UCM951.Class import UCMULD
        UCMULDTmp = UCMULD(Type, No, Owner)  # 创建UCMULD对象
        from ReadTXT.UCM951.Variable import UCMULDLst
        UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表
    else:  # 4位号
        No = '0' + item[16:20]  # 4位号
        Owner = item[20:22]  # 所有人
        from ReadTXT.UCM951.Class import UCMULD
        UCMULDTmp = UCMULD(Type, No, Owner)  # 创建UCMULD对象
        from ReadTXT.UCM951.Variable import UCMULDLst
        UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表