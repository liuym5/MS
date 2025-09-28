def ReadCPM(Path):  # 读取CPM,返回是否有叠板
    from ReadTXT.Function import ReadTXT
    CPM = ReadTXT(Path)  # 返回CPM文件文本
    CPM = CPM.split('-')  # 以-分割字符串
    TypeTup = ('AKE', 'PMC', 'PAG', 'PLA', 'PAJ', 'P1P')  # 类型元组
    StackTF = False  # 无叠板
    for item in CPM:  # 遍历CPM字符串列表
        Type, i = FindType(TypeTup, item)  # 返回类型,类型首字母下标
        Content = FindContent(item)  # 返回内容
        if Type in ('PMC', 'PAG', 'PLA', 'PAJ', 'P1P') and Content == 'X':  # 有叠板
            StackTF = True  # 有叠板
        if i > -1:  # 找到类型
            No = item[i + 3:i + 7]  # 4位号
            Owner = item[i + 7:i + 9]  # 所有人
            if No.isdigit() and Owner in ('MS', 'R7', 'R9', 'CI', 'C6', 'K8', 'KE'):  # 4位号为数字并且所有人为MS或R7或R9或CI或C6或K8或KE
                No = '0' + No  # 号前面加个0补充到5位
                from ReadTXT.CPM.Class import CPMULD
                CPMULDTmp = CPMULD(Type, No, Owner, Content)  # 创建CPMULD对象
                from ReadTXT.CPM.Variable import CPMULDLst
                CPMULDLst.append(CPMULDTmp)  # 添加CPMULD对象到CPMULD对象列表
            No = item[i + 3:i + 8]  # 5位号
            Owner = item[i + 8:i + 10]  # 所有人
            if No.isdigit() and Owner in ('MS', 'R7', 'R9', 'CI', 'C6', 'K8', 'KE'):  # 5位号为数字并且所有人为MS或R7或R9或CI或C6或K8或KE
                from ReadTXT.CPM.Class import CPMULD
                CPMULDTmp = CPMULD(Type, No, Owner, Content)  # 创建CPMULD对象
                from ReadTXT.CPM.Variable import CPMULDLst
                CPMULDLst.append(CPMULDTmp)  # 添加CPMULD对象到CPMULD对象列表
    import operator
    cmpfun = operator.attrgetter('No')  # 参数为排序依据的属性，可以有多个，这里优先id，使用时按需求改换参数即可
    from ReadTXT.CPM.Variable import CPMULDLst
    CPMULDLst.sort(key=cmpfun)  # 使用时改变列表名即可
    return StackTF  # 返回是否有叠板

def FindType(TypeTup, item):  # 返回类型,类型首字母下标或返回空字符串,-1
    i = 0  # 类型首字母下标
    for type in TypeTup:  # 遍历元组
        i = item.find(type)  # 得到类型首字母下标
        if i > -1:  # 找到类型
            return type, i  # 返回类型,类型首字母下标
    return '', -1  # 返回空字符串,-1

def FindContent(item):  # 返回内容
    if item.rfind('/B') > -1:  # 右找到/B
        return 'B' # 返回B
    if item.rfind('/X') > -1:  # 右找到/X
        return 'X'  # 返回X
    if item.rfind('/C') > -1:  # 右找到/C
        return 'C' # 返回C
    if item.rfind('/L') > -1:  # 右找到/L
        return 'L'  # 返回L

def ReadCPMULD():  # 读取CPMULDLst
    from ReadTXT.CPM.Variable import CPMULDLst
    for cpmuld in CPMULDLst:  # 遍历CPMULDLst
        Type = cpmuld.Type  # 类型
        No = cpmuld.No  # 号
        Owner = cpmuld.Owner  # 所有人
        from ReadTXT.UCM951.Class import UCMULD
        UCMULDTmp = UCMULD(Type, No, Owner)  # 创建UCMULD对象
        from ReadTXT.UCM951.Variable import UCMULDLst
        UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表