def ReadCPM(CPM):  # 读取CPM
    CPM = CPM.split('-')  # 以-分割字符串
    TypeTup = ('AKE', 'PMC', 'PLA', 'PAG', 'P1P')  # 类型元组
    Content = ''  # 内容
    No = ''  # 号
    Owner = ''  # 所有人
    for Pos in CPM:  # 遍历CPM字符串列表
        Type, i = FindType(TypeTup, Pos)  # 返回类型,类型首字母下标
        Content = FindContent(Pos)  # 返回内容
        if i > -1:  # 找到类型
            for i in range(3, len(Pos)):  # 遍历Pos字符串
                if Pos[i].isdigit():  # 是数字
                    No = No + Pos[i]  # 拼接号
                if Pos[i - 1] + Pos[i] in ('MS', 'DS'):  # 2个字符是MS或DS
                    Owner = Pos[i - 1] + Pos[i]  # 拼接所有人
                if Pos[i] == '/':  # 是斜线
                    if len(No) == 5:  # 号为5位
                        from ReadPDF.CPM.Class import CPMULD
                        CPMULDTmp = CPMULD(Type, No, Owner, Content)  # 创建CPMULD对象
                        from ReadPDF.CPM.Variable import CPMULDLst
                        CPMULDLst.append(CPMULDTmp)  # 添加CPMULD对象到CPMULD对象列表
                    elif len(No) == 4 and Pos[i - 1] == 'S':  # 号为4位并且左边为S
                        No = '0' + No  # 号前面加个0补充到5位
                        from ReadPDF.CPM.Class import CPMULD
                        CPMULDTmp = CPMULD(Type, No, Owner, Content)  # 创建CPMULD对象
                        from ReadPDF.CPM.Variable import CPMULDLst
                        CPMULDLst.append(CPMULDTmp)  # 添加CPMULD对象到CPMULD对象列表
                    No = ''  # 恢复号为默认值空

def FindType(TypeTup, Pos):  # 返回类型,类型首字母下标或返回空字符串,-1
    i = 0  # 类型首字母下标
    for Type in TypeTup:  # 遍历元组
        i = Pos.find(Type)  # 得到类型首字母下标
        if i > -1:  # 找到类型
            return Type, i  # 返回类型,类型首字母下标
    return '', -1  # 返回空字符串,-1

def FindContent(Pos):  # 返回内容
    if Pos.find('/B') > -1:  # 找到/B
        return 'B' # 返回B
    if Pos.find('/C') > -1:  # 找到/C
        return 'C' # 返回C
    if Pos.find('/X') > -1:  # 找到/X
        return 'X' # 返回X