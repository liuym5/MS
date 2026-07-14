def WritFlight(Path, Date, Flight):  # 写Flights文本文件
    Txt = ('站点: PVG\n'
           '航班号/航班日期：MS952/')
    import datetime
    DateDT = datetime.datetime.strptime(Date, '%y%m%d')  # 日期字符串转日期格式
    Date = DateDT.strftime('%d%b').upper()  # 日期格式转日期字符串并大写
    Txt = (Txt + Date + '\n'
           '实际起飞时间： 23:55（延误0min， 原因：不详）\n'
           '机型：' + Flight.ACType + '\n'
           '旅客人数：' + Flight.PAX + '人\n'
           '旅客使用板箱：')
    Txt = Txt + GetLULD(Flight)  # 拼接返回行李集装器字符串
    MULD = GetMULD(Flight)  # 返回MCO集装器字符串
    PULD = GetPULD(Flight)  # 返回预报MCO集装器字符串
    if MULD != '':  # 有MCO集装器
        Txt = Txt + '(其中MCO, ' + MULD + ', ' + Flight.MDest.replace('/', ' ')
        if PULD != MULD:  # 和预报MCO有差异
            Txt = PreMCO(Txt, DateDT, PULD)  # 得到组织好的PMCO文本
        else:  # 和预报MCO一致
            Txt = Txt + ')\n'
    else:  # 无MCO集装器
        Txt = Txt + '(无MCO'
        if PULD != MULD:  # 和预报MCO有差异
            Txt = PreMCO(Txt, DateDT, PULD)  # 得到组织好的PMCO文本
        else:  # 和预报MCO一致
            Txt = Txt + ')\n'
    Txt = Txt + '货运使用板箱：'
    Txt = Txt + GetCULD(Flight) + '\n'  # 拼接返回货集装器字符串
    Txt = (Txt + '走货重量：GW ' + Flight.GW + ' KG CW ' + Flight.CW + ' KG\n'
           '拉货：')
    Txt = Txt + GetOULD(Flight) + '\n'  # 拼接返回拉货集装器字符串
    Txt = (Txt + '未使用板箱情况：')
    Txt = Txt + GetRULD(Flight)  # 拼接返回空集装器字符串
    Txt = Txt + GetORsn(Flight.ORsn)  # 拼接返回拉货原因符串
    from WritTXT.Function import WritTXT
    WritTXT(Path, 'w+', Txt)  # 写TXT文件

def GetLULD(Flight):  # 返回行李集装器字符串
    ULDLst = []  # 集装器列表为空
    if Flight.LPMC != '':  # 有行李PMC
        ULDLst.append(Flight.LPMC + 'PMC')  # 添加行李PMC到集装器列表
    if Flight.LPAG != '':  # 有行李PAG
        ULDLst.append(Flight.LPAG + 'PAG')  # 添加行李PAG到集装器列表
    if Flight.LPLA != '':  # 有行李PLA
        ULDLst.append(Flight.LPLA + 'PLA')  # 添加行李PLA到集装器列表
    if Flight.LAKE != '':  # 有行李AKE
        ULDLst.append(Flight.LAKE + 'AKE')  # 添加行李AKE到集装器列表
    Len = len(ULDLst)  # 集装器列表长度
    ULD = ''  # 集装器为空
    for i in range(Len):  # 遍历列表
        ULD = ULD + ULDLst[i]  # 添加集装器字符串
        if i + 1 < Len:  # 不是列表最后
            ULD = ULD + '+'  # 末尾添加加号
    if ULD == '':  # 无行李集装器
        ULD = '无'
    return ULD  # 返回集装器字符串

def GetMULD(Flight):  # 返回MCO集装器字符串
    ULDLst = []  # 集装器列表为空
    if Flight.MPMC != '':  # 有MCOPMC
        ULDLst.append(Flight.MPMC + 'PMC')  # 添加MCOPMC到集装器列表
    if Flight.MPAG != '':  # 有MCOPAG
        ULDLst.append(Flight.MPAG + 'PAG')  # 添加MCOPAG到集装器列表
    if Flight.MPLA != '':  # 有MCOPLA
        ULDLst.append(Flight.MPLA + 'PLA')  # 添加MCOPLA到集装器列表
    if Flight.MAKE != '':  # 有MCOAKE
        ULDLst.append(Flight.MAKE + 'AKE')  # 添加MCOAKE到集装器列表
    Len = len(ULDLst)  # 集装器列表长度
    ULD = ''  # 集装器为空
    for i in range(Len):  # 遍历列表
        ULD = ULD + ULDLst[i]  # 添加集装器字符串
        if i + 1 < Len:  # 不是列表最后
            ULD = ULD + '+'  # 末尾添加加号
    return ULD  # 返回集装器字符串

def GetPULD(Flight):  # 返回预报MCO集装器字符串
    ULDLst = []  # 集装器列表为空
    if Flight.PPMC != '':  # 有预报MCOPMC
        ULDLst.append(Flight.PPMC + 'PMC')  # 添加预报MCOPMC到集装器列表
    if Flight.PPAG != '':  # 有预报MCOPAG
        ULDLst.append(Flight.PPAG + 'PAG')  # 添加预报MCOPAG到集装器列表
    Len = len(ULDLst)  # 集装器列表长度
    ULD = ''  # 集装器为空
    for i in range(Len):  # 遍历列表
        ULD = ULD + ULDLst[i]  # 添加集装器字符串
        if i + 1 < Len:  # 不是列表最后
            ULD = ULD + '+'  # 末尾添加加号
    return ULD  # 返回集装器字符串

def PreMCO(Txt, DateDT, PULD):  # 返回组织好的预报MCO集装器文本
    Txt = Txt + ', 但实际客运'
    import datetime
    DateDT = DateDT + datetime.timedelta(days=-1)  # 日期减1天
    Date = DateDT.strftime('%d%b').upper()  # 日期格式转日期字符串并大写
    return Txt + Date + '预报' + PULD + ')\n'  # 返回组织好的PMCO文本

def GetCULD(Flight):  # 返回货集装器字符串
    ULDLst = []  # 集装器列表为空
    if Flight.CPMC != '':  # 有货PMC
        ULDLst.append(Flight.CPMC + 'PMC')  # 添加货PMC到集装器列表
    if Flight.CPAG != '':  # 有货PAG
        ULDLst.append(Flight.CPAG + 'PAG')  # 添加货PAG到集装器列表
    if Flight.CPLA != '':  # 有货PLA
        ULDLst.append(Flight.CPLA + 'PLA')  # 添加货PLA到集装器列表
    if Flight.CAKE != '':  # 有货AKE
        ULDLst.append(Flight.CAKE + 'AKE')  # 添加货AKE到集装器列表
    Len = len(ULDLst)  # 集装器列表长度
    ULD = ''  # 集装器为空
    for i in range(Len):  # 遍历列表
        ULD = ULD + ULDLst[i]  # 添加集装器字符串
        if i + 1 < Len:  # 不是列表最后
            ULD = ULD + '+'  # 末尾添加加号
    if ULD == '':  # 无集装器
        ULD = '无'
    return ULD  # 返回集装器字符串

def GetOULD(Flight):  # 返回拉货集装器字符串
    ULDLst = []  # 集装器列表为空
    if Flight.OPMC != '':  # 有拉货PMC
        ULDLst.append(Flight.OPMC + 'PMC')  # 添加拉货PMC到集装器列表
    if Flight.OPAG != '':  # 有拉货PAG
        ULDLst.append(Flight.OPAG + 'PAG')  # 添加拉货PAG到集装器列表
    if Flight.OPLA != '':  # 有拉货PLA
        ULDLst.append(Flight.OPLA + 'PLA')  # 添加拉货PLA到集装器列表
    if Flight.OAKE != '':  # 有拉货AKE
        ULDLst.append(Flight.OAKE + 'AKE')  # 添加拉货AKE到集装器列表
    Len = len(ULDLst)  # 集装器列表长度
    ULD = ''  # 集装器为空
    for i in range(Len):  # 遍历列表
        ULD = ULD + ULDLst[i]  # 添加集装器字符串
        if i + 1 < Len:  # 不是列表最后
            ULD = ULD + '+'  # 末尾添加加号
    if ULD == '':  # 无集装器
        ULD = '无'
    return ULD  # 返回集装器字符串

def GetRULD(Flight):  # 返回空集装器字符串
    ULDLst = []  # 集装器列表为空
    if Flight.RPMC != '':  # 有空PMC
        ULDLst.append(Flight.RPMC + 'PMC')  # 添加空PMC到集装器列表
    if Flight.RPAG != '':  # 有空PAG
        ULDLst.append(Flight.RPAG + 'PAG')  # 添加空PAG到集装器列表
    if Flight.RPLA != '':  # 有空PLA
        ULDLst.append(Flight.RPLA + 'PLA')  # 添加空PLA到集装器列表
    if Flight.RAKE != '':  # 有空AKE
        ULDLst.append(Flight.RAKE + 'AKE')  # 添加空AKE到集装器列表
    Len = len(ULDLst)  # 集装器列表长度
    ULD = ''  # 集装器为空
    for i in range(Len):  # 遍历列表
        ULD = ULD + ULDLst[i]  # 添加集装器字符串
        if i + 1 < Len:  # 不是列表最后
            ULD = ULD + '+'  # 末尾添加加号
    if ULD == '':  # 无集装器
        ULD = '无'
    return ULD  # 返回集装器字符串

def GetORsn(ORsn):  # 返回拉货原因字符串
    if ORsn == 'P':  # 限载
        return ' 限载'
    if ORsn == 'B':  # 限平衡
        return ' 平衡'
    if ORsn == 'A':  # 限飞机
        return ' 飞机问题'
    return ''