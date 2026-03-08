def WritFlight(Path, Date, Flight):  # 写Flights文本文件
    Txt = ('站点: PVG\n'
           '航班号/航班日期：MS952/' + Date + '\n'
           '实际起飞时间： 23:55（延误0min， 原因：不详）\n'
           '机型：')
    if Flight.ACType == 'B787-900':  # 机型为787-900
        ACType = 'B789'
    else:  # 机型为777-300
        ACType = 'B773'
    Txt = (Txt + ACType + '\n'
           '旅客人数：' + Flight.PAX + '人\n'
           '旅客使用板箱：')
    ULDLst = []  # 行李集装器列表为空
    if Flight.LPMC != '':  # 有行李PMC
        ULDLst.append(Flight.LPMC + 'PMC')  # 添加行李PMC到行李集装器列表
    if Flight.LPAG != '':  # 有行李PAG
        ULDLst.append(Flight.LPAG + 'PAG')  # 添加行李PAG到行李集装器列表
    if Flight.LPLA != '':  # 有行李PLA
        ULDLst.append(Flight.LPLA + 'PLA')  # 添加行李PLA到行李集装器列表
    if Flight.LAKE != '':  # 有行李AKE
        ULDLst.append(Flight.LAKE + 'AKE')  # 添加行李AKE到行李集装器列表
    Len = len(ULDLst)  # 行李集装器列表长度
    ULD = ''  # 行李集装器为空
    for i in range(Len):  # 遍历列表
        ULD = ULD + ULDLst[i]  # 添加行李集装器字符串
        if i + 1 < Len:  # 不是列表最后
            ULD = ULD + '+'  # 末尾添加加号
    if ULD == '':  # 无行李集装器
        ULD = '无'
    Txt = Txt + ULD
    ULDLst = []  # MCO集装器列表为空
    if Flight.MPMC != '':  # 有MCOPMC
        ULDLst.append(Flight.MPMC + 'PMC')  # 添加MCOPMC到MCO集装器列表
    if Flight.MPAG != '':  # 有MCOPAG
        ULDLst.append(Flight.MPAG + 'PAG')  # 添加MCOPAG到MCO集装器列表
    if Flight.MPLA != '':  # 有MCOPLA
        ULDLst.append(Flight.MPLA + 'PLA')  # 添加MCOPLA到MCO集装器列表
    if Flight.MAKE != '':  # 有MCOAKE
        ULDLst.append(Flight.MAKE + 'AKE')  # 添加MCOAKE到MCO集装器列表
    Len = len(ULDLst)  # MCO集装器列表长度
    ULD = ''  # MCO集装器为空
    for i in range(Len):  # 遍历列表
        ULD = ULD + ULDLst[i]  # 添加MCO集装器字符串
        if i + 1 < Len:  # 不是列表最后
            ULD = ULD + '+'  # 末尾添加加号
    if ULD != '':  # 有MCO集装器
        Txt = Txt + '(其中MCO, ' + ULD + ', ' + Flight.MDest.replace('/', ' ') + ')\n'
    else:  # 无MCO集装器
        Txt = Txt + '(无MCO)\n'
    Txt = Txt + '货运使用板箱：'
    ULDLst = []  # 货集装器列表为空
    if Flight.CPMC != '':  # 有货PMC
        ULDLst.append(Flight.CPMC + 'PMC')  # 添加货PMC到货集装器列表
    if Flight.CPAG != '':  # 有货PAG
        ULDLst.append(Flight.CPAG + 'PAG')  # 添加货PAG到货集装器列表
    if Flight.CPLA != '':  # 有货PLA
        ULDLst.append(Flight.CPLA + 'PLA')  # 添加货PLA到货集装器列表
    if Flight.CAKE != '':  # 有货AKE
        ULDLst.append(Flight.CAKE + 'AKE')  # 添加货AKE到货集装器列表
    Len = len(ULDLst)  # 货集装器列表长度
    ULD = ''  # 货集装器为空
    for i in range(Len):  # 遍历列表
        ULD = ULD + ULDLst[i]  # 添加货集装器字符串
        if i + 1 < Len:  # 不是列表最后
            ULD = ULD + '+'  # 末尾添加加号
    if ULD == '':  # 无货集装器
        ULD = '无'
    Txt = (Txt + ULD + '\n'
           '走货重量：GW ' + Flight.GW + ' KG CW ' + Flight.CW + ' KG\n'
           '拉货：')
    ULDLst = []  # 拉货集装器列表为空
    if Flight.OPMC != '':  # 有拉货PMC
        ULDLst.append(Flight.OPMC + 'PMC')  # 添加拉货PMC到拉货集装器列表
    if Flight.OPAG != '':  # 有拉货PAG
        ULDLst.append(Flight.OPAG + 'PAG')  # 添加拉货PAG到拉货集装器列表
    if Flight.OPLA != '':  # 有拉货PLA
        ULDLst.append(Flight.OPMC + 'PLA')  # 添加拉货PLA到拉货集装器列表
    if Flight.OAKE != '':  # 有拉货AKE
        ULDLst.append(Flight.OAKE + 'AKE')  # 添加拉货AKE到拉货集装器列表
    Len = len(ULDLst)  # 拉货集装器列表长度
    ULD = ''  # 拉货集装器为空
    for i in range(Len):  # 遍历列表
        ULD = ULD + ULDLst[i]  # 添加拉货集装器字符串
        if i + 1 < Len:  # 不是列表最后
            ULD = ULD + '+'  # 末尾添加加号
    if ULD == '':  # 无拉货集装器
        ULD = '无'
    Txt = (Txt + ULD + '\n'
           '未使用板箱情况：')
    ULDLst = []  # 空集装器列表为空
    if Flight.RPMC != '':  # 有空PMC
        ULDLst.append(Flight.RPMC + 'PMC')  # 添加空PMC到空集装器列表
    if Flight.RPAG != '':  # 有空PAG
        ULDLst.append(Flight.RPAG + 'PAG')  # 添加空PAG到空集装器列表
    if Flight.RPLA != '':  # 有空PLA
        ULDLst.append(Flight.RPLA + 'PLA')  # 添加空PLA到空集装器列表
    if Flight.RAKE != '':  # 有空AKE
        ULDLst.append(Flight.RAKE + 'AKE')  # 添加空AKE到空集装器列表
    Len = len(ULDLst)  # 空集装器列表长度
    ULD = ''  # 空集装器为空
    for i in range(Len):  # 遍历列表
        ULD = ULD + ULDLst[i]  # 添加空集装器字符串
        if i + 1 < Len:  # 不是列表最后
            ULD = ULD + '+'  # 末尾添加加号
    if ULD == '':  # 无空集装器
        ULD = '无'
    Txt = Txt + ULD
    if Flight.ORsn == 'P':  # 限载
        Txt = Txt + ' 限载'
    Txt = (Txt + '\n'
           '\n')
    TXT = open(Path, 'a', encoding='utf-8')  # 返回文本文件对象
    TXT.write(Txt)  # 写文本文件
    TXT.close()  # 关闭文本文件对象