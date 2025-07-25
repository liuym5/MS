def ReadAKE952(Path):  # 读取AKE952
    from ReadTXT.Function import ReadTXT
    AKE = ReadTXT(Path)  # 返回AKE952文件文本
    AKE = AKE.split('\n')  # 以换行符分割字符串
    for item in AKE:  # 遍历AKE字符串列表
        from ReadTXT.CPM.Class import CPMULD
        CPMULDTmp = CPMULD('AKE', item, 'MS', 'B')  # 创建CPMULD对象
        from ReadTXT.CPM.Variable import CPMULDLst
        CPMULDLst.append(CPMULDTmp)  # 添加CPMULD对象到CPMULD对象列表
    import operator
    cmpfun = operator.attrgetter('No')  # 参数为排序依据的属性，可以有多个，这里优先id，使用时按需求改换参数即可
    from ReadTXT.CPM.Variable import CPMULDLst
    CPMULDLst.sort(key=cmpfun)  # 使用时改变列表名即可