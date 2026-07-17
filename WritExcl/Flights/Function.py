from ReadExcl.Flight.Class import Flight


def WritACType(Path, Flight):  # 写PRELOAD表格文件机号
    import win32com.client
    XL = win32com.client.Dispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    WB = XL.Workbooks.Open(Path)  # 返回Statistic表格对象
    ST = WB.Worksheets('PRE-LOAD')  # 返回当月当年页对象
    ST.Cells(3, 3).Value = Flight.ACType  # 写机号
    WB.Save()  # 保存Statistic表格
    WB.Close()  # 关闭Statistic表格对象
    XL.Quit()  # 关闭Excel

def WritStatistic(Path, SN, r, Flight):  # 写Statistic表格文件
    import win32com.client
    XL = win32com.client.Dispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    WB = XL.Workbooks.Open(Path)  # 返回Statistic表格对象
    ST = WB.Worksheets(SN)  # 返回当月当年页对象
    if Flight.GW != '0':  # 有货
        ST.Cells(r, 2).Value = Flight.GW  # 写重量
    else:  # 无货
        ST.Cells(r, 2).Value = 'NIL'  # 写重量为NIL
    CPMC = 0  # 货PMC为0
    CPAG = 0  # 货PAG为0
    CPLA = 0  # 货PLA为0
    if Flight.CPMC != '':  # 有货PMC
        CPMC = int(Flight.CPMC)  # 得到货PMC
    if Flight.CPAG != '':  # 有货PAG
        CPAG = int(Flight.CPAG)  # 得到货PAG
    if Flight.CPLA != '':  # 有货PLA
        CPLA = int(Flight.CPLA)  # 得到货PMC
    ST.Cells(r, 12).Value = (CPMC + CPAG + CPLA) * 2  # 写雨布张数
    WB.Save()  # 保存Statistic表格
    WB.Close()  # 关闭Statistic表格对象
    XL.Quit()  # 关闭Excel

def WritWaterproof(Path, SN, r, Flight):  # 写雨布表格文件
    import win32com.client
    XL = win32com.client.Dispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    WB = XL.Workbooks.Open(Path)  # 返回Statistic表格对象
    ST = WB.Worksheets(SN)  # 返回当月当年页对象
    ULDLst = []  # 货集装器列表
    CPMC = 0  # 货PMC为0
    CPAG = 0  # 货PAG为0
    CPLA = 0  # 货PLA为0
    if Flight.CPMC != '':  # 有货PMC
        ULDLst.append(Flight.CPMC + 'PMC')  # 添加货PMC到货集装器列表
        CPMC = int(Flight.CPMC)  # 得到货PMC
    if Flight.CPAG != '':  # 有货PAG
        ULDLst.append(Flight.CPAG + 'PAG')  # 添加货PAG到货集装器列表
        CPAG = int(Flight.CPAG)  # 得到货PAG
    if Flight.CPLA != '':  # 有货PLA
        ULDLst.append(Flight.CPLA + 'PLA')  # 添加货PLA到货集装器列表
        CPLA = int(Flight.CPLA)  # 得到货PLA
    Len = len(ULDLst)  # 集装器列表长度
    ULD = ''  # 货集装器为空
    for i in range(Len):  # 遍历列表
        ULD = ULD + ULDLst[i]  # 添加货集装器字符串
        if i + 1 < Len:  # 不是列表最后
            ULD = ULD + '+'  # 末尾添加加号
    if ULD == '':  # 无货
        ULD = 'NIL'  # 无货集装器
    ST.Cells(r, 4).Value = ULD  # 写货集装器
    ST.Cells(r, 5).Value = (CPMC + CPAG + CPLA) * 2  # 写雨布张数
    WB.Save()  # 保存Statistic表格
    WB.Close()  # 关闭Statistic表格对象
    XL.Quit()  # 关闭Excel

def WritMCO(Path, SN, r, Flight):  # 写MCO表格文件
    import win32com.client
    XL = win32com.client.Dispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    WB = XL.Workbooks.Open(Path)  # 返回Statistic表格对象
    ST = WB.Worksheets(SN)  # 返回当月当年页对象
    ST.Cells(r, 2).Value = GetMCONo(ST, r)  # 写返回MCO文件序号数字
    ST.Cells(r, 4).Value = Flight.ACType  # 写机型
    ST.Cells(r, 5).Value = Flight.GW  # 写重量
    ST.Cells(r, 6).Value = Flight.CPMC  # 写货PMC
    ST.Cells(r, 7).Value = Flight.CPAG  # 写货PAG
    ST.Cells(r, 8).Value = Flight.CPLA  # 写货PLA
    ST.Cells(r, 9).Value = Flight.CAKE  # 写货AKE
    ST.Cells(r, 10).Value = Flight.MDest  # 写MCO目的地
    ST.Cells(r, 11).Value = Flight.MPcs  # 写MCO件数
    ST.Cells(r, 12).Value = Flight.MGW  # 写MCO重量
    ST.Cells(r, 13).Value = Flight.MPMC  # 写MCOPMC
    ST.Cells(r, 14).Value = Flight.MPAG  # 写MCOPAG
    ST.Cells(r, 15).Value = Flight.MPLA  # 写MCOPLA
    ST.Cells(r, 16).Value = Flight.MAKE  # 写MCOAKE
    ST.Cells(r, 17).Value = Flight.RPMC  # 写空PMC
    ST.Cells(r, 18).Value = Flight.RPAG  # 写空PAG
    ST.Cells(r, 19).Value = Flight.RPLA  # 写空PLA
    ST.Cells(r, 20).Value = Flight.RAKE  # 写空AKE
    ST.Cells(r, 21).Value = Flight.OPMC  # 写拉货PMC
    ST.Cells(r, 22).Value = Flight.OPAG  # 写拉货PAG
    ST.Cells(r, 23).Value = Flight.OPLA  # 写拉货PLA
    ST.Cells(r, 24).Value = Flight.OAKE  # 写拉货AKE
    ST.Cells(r, 25).Value = Flight.OGW  # 写拉货重量
    WB.Save()  # 保存Statistic表格
    WB.Close()  # 关闭Statistic表格对象
    XL.Quit()  # 关闭Excel

def WritMonitor(Path, SN, r, Flight):  # 写Monitor表格文件
    import win32com.client
    XL = win32com.client.Dispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    WB = XL.Workbooks.Open(Path)  # 返回Statistic表格对象
    ST = WB.Worksheets(SN)  # 返回当月当年页对象
    ST.Cells(r, 1).Value = GetMonitorNo(ST, r)  # 写返回序号数字
    ST.Cells(r, 6).Value = Flight.ACType  # 写机型
    ST.Cells(r, 7).Value = GetCW(Flight)  # 写返回计费重量字符串
    ST.Cells(r, 10).Value = Flight.CPMC  # 写货PMC
    ST.Cells(r, 11).Value = Flight.CPAG  # 写货PAG
    ST.Cells(r, 12).Value = Flight.CPLA  # 写货PLA
    ST.Cells(r, 13).Value = Flight.CAKE  # 写货AKE
    ST.Cells(r, 14).Value = Flight.LPMC  # 写行李PMC
    ST.Cells(r, 15).Value = Flight.LPAG  # 写行李PAG
    ST.Cells(r, 16).Value = Flight.LPLA  # 写行李PLA
    ST.Cells(r, 17).Value = Flight.LAKE  # 写行李AKE
    ST.Cells(r, 18).Value = Flight.OGW  # 写拉货重量
    ST.Cells(r, 19).Value = Flight.OPMC  # 写拉货PMC
    ST.Cells(r, 20).Value = Flight.OPAG  # 写拉货PAG
    ST.Cells(r, 21).Value = Flight.OPLA  # 写拉货PLA
    ST.Cells(r, 22).Value = Flight.OAKE  # 写拉货AKE
    ORsn = GetORsn(Flight.ORsn)  # 返回拉货原因字符串
    ST.Cells(r, 23).Value = ORsn  # 写拉货原因
    ST.Cells(r, 24).Value = GetMCOULD(Flight)  # 写返回MCO集装器字符串
    ST.Cells(r, 25).Value = Flight.PAX  # 写人数
    RULD = GetRULD(Flight)  # 返回空集装器字符串
    ST.Cells(r, 26).Value = RULD  # 写空集装器字符串
    ST.Cells(r, 27).Value = GetRRsn(ORsn, RULD, Flight)  # 写返回空舱位原因字符串
    ST.Cells(r, 28).Value = Flight.ULoad  # 写剩余载量
    ST.Cells(r, 29).Value = Flight.ACNo  # 写机号
    ST.Cells(r, 30).Value = Flight.Load  # 写载重
    ST.Cells(r, 31).Value = Flight.TOW  # 写起飞重量
    WB.Save()  # 保存Statistic表格
    WB.Close()  # 关闭Statistic表格对象
    XL.Quit()  # 关闭Excel

def WritMonitor2(Path, SN, r, Flight):  # 写Monitor副本表格文件
    import win32com.client
    XL = win32com.client.Dispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    WB = XL.Workbooks.Open(Path)  # 返回Statistic表格对象
    ST = WB.Worksheets(SN)  # 返回当月当年页对象
    ST.Cells(r, 2).Value = Flight.Date  # 写日期
    ST.Cells(r, 6).Value = Flight.ACType  # 写机型
    ST.Cells(r, 7).Value = GetCW(Flight)  # 写返回计费重量字符串
    ST.Cells(r, 10).Value = Flight.CPMC  # 写货PMC
    ST.Cells(r, 11).Value = Flight.CPAG  # 写货PAG
    ST.Cells(r, 12).Value = Flight.CPLA  # 写货PLA
    ST.Cells(r, 13).Value = Flight.CAKE  # 写货AKE
    ST.Cells(r, 14).Value = Flight.LPMC  # 写行李PMC
    ST.Cells(r, 15).Value = Flight.LPAG  # 写行李PAG
    ST.Cells(r, 16).Value = Flight.LPLA  # 写行李PLA
    ST.Cells(r, 17).Value = Flight.LAKE  # 写行李AKE
    ST.Cells(r, 18).Value = Flight.OGW  # 写拉货重量
    ST.Cells(r, 19).Value = Flight.OPMC  # 写拉货PMC
    ST.Cells(r, 20).Value = Flight.OPAG  # 写拉货PAG
    ST.Cells(r, 21).Value = Flight.OPLA  # 写拉货PLA
    ST.Cells(r, 22).Value = Flight.OAKE  # 写拉货AKE
    ORsn = GetORsn(Flight.ORsn)  # 返回拉货原因字符串
    ST.Cells(r, 23).Value = ORsn  # 写拉货原因
    ST.Cells(r, 24).Value = GetMCOULD(Flight)  # 写返回MCO集装器字符串
    ST.Cells(r, 25).Value = Flight.PAX  # 写人数
    RULD = GetRULD(Flight)  # 返回空集装器字符串
    ST.Cells(r, 26).Value = RULD  # 写空集装器字符串
    ST.Cells(r, 27).Value = GetRRsn(ORsn, RULD, Flight)  # 写返回空舱位原因字符串
    WB.Save()  # 保存Statistic表格
    WB.Close()  # 关闭Statistic表格对象
    XL.Quit()  # 关闭Excel

def WritVerify(Path, SN, r, Flight):  # 写对账表格文件
    import win32com.client
    XL = win32com.client.Dispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    WB = XL.Workbooks.Open(Path)  # 返回Statistic表格对象
    ST = WB.Worksheets(SN)  # 返回当月当年页对象
    if ST.Cells(r, 2).Text == 'MS0951':  # 进港航班
        r += 1  # 行号加1
    ST.Cells(r, 7).Value = Flight.GW  # 写重量
    WB.Save()  # 保存Statistic表格
    WB.Close()  # 关闭Statistic表格对象
    XL.Quit()  # 关闭Excel

def GetMCONo(ST, r):  # 返回MCO文件序号数字
    No = ST.Cells(r - 1, 2).Text  # 上1行序号
    if No.isdigit():  # 是数字
        return int(No) + 1  # 序号加1
    return 1  # 序号加1

def GetMonitorNo(ST, r):  # 返回Monitor文件序号数字
    No = ST.Cells(r - 1, 1).Text  # 上1行序号
    if No.isdigit():  # 是数字
        return int(No) + 1  # 序号加1
    No = ST.Cells(r-2, 1).Text  # 上2行序号
    return int(No) + 1  # 序号加1

def GetCW(Flight):  # 返回计费重量字符串
    if Flight.CW != '0':  # 有货
        return Flight.CW  # 写计费重量字符串
    return 'NIL'  # 返回计费重量为NIL字符串

def GetORsn(ORsn):  # 返回拉货原因
    if ORsn == 'P':  # 限载
        return 'Payload restriction'
    if ORsn == 'S':  # 限舱位
        return 'Lack of space'
    if ORsn == 'B':  # 限平衡
        return 'Balance problem'
    if ORsn == 'A':  # 限飞机
        return 'Aircraft problem'
    return ''

def GetMCOULD(Flight):  # 返回MCO集装器字符串
    ULDLst = []  # 集装器列表
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

def GetRULD(Flight):  # 返回空集装器字符串
    ULDLst = []  # 集装器列表
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
    return ULD  # 返回集装器字符串

def GetRRsn(ORsn, RULD, Flight):  # 返回空舱位原因字符串
    if RULD in ['', '1AKE', '2AKE']:  # 集装器为空或1AKE或2AKE
        return ''  # 返回空字符串
    if ifNoshowMCO(Flight):  # 是否MCO no show
        return 'MCO no show'
    if Flight.OGW == '':  # 无拉货
        return 'No cargo'  # 返回无货字符串
    return ORsn  # 返回拉货原因字符串

def ifNoshowMCO(Flight):  # 是否MCO no show
    PPMC = 0  # 预计MCOPMC为0
    PPAG = 0  # 预计MCOPAG为0
    if Flight.PPMC != '':  # 有预计MCOPMC
        PPMC = int(Flight.PPMC)  # 得到预计MCOPMC数量
    if Flight.PPAG != '':  # 有预计MCOPAG
        PPAG = int(Flight.PPAG)  # 得到预计MCOPAG数量
    MPMC = 0  # MCOPMC为0
    MPAG = 0  # MCOPAG为0
    if Flight.MPMC != '':  # 有MCOPMC
        MPMC = int(Flight.MPMC)  # 得到MCOPMC数量
    if Flight.MPAG != '':  # 有MCOPAG
        MPAG = int(Flight.MPAG)  # 得到MCOPAG数量
    if PPMC + PPAG > MPMC + MPAG:  # MCO no show
        return True
    return False