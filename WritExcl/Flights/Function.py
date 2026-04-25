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
    No = ST.Cells(r-1, 2).Text  # 上1行序号
    if No.isdigit():  # 是数字
        No = int(No) + 1  # 序号加1
    else:  # 不是数字
        No = 1  # 序号为1
    ST.Cells(r, 2).Value = No  # 写序号
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
    No = ST.Cells(r-1, 1).Text  # 上1行序号
    if No.isdigit():  # 是数字
        No = int(No) + 1  # 序号加1
    else:  # 不是数字
        No = ST.Cells(r-2, 1).Text  # 上2行序号
        No = int(No) + 1  # 序号加1
    ST.Cells(r, 1).Value = No  # 写序号
    ST.Cells(r, 6).Value = Flight.ACType  # 写机型
    if Flight.CW != '0':  # 有货
        ST.Cells(r, 7).Value = Flight.CW  # 写计费重量
    else:  # 无货
        ST.Cells(r, 7).Value = 'NIL'  # 写计费重量为NIL
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
    if Flight.ORsn == 'P':  # 限载
        ST.Cells(r, 23).Value = 'Payload restriction'  # 写载量限制
    elif Flight.ORsn == 'S':  # 限舱位
        ST.Cells(r, 23).Value = 'Lack of space'  # 写舱位限制
    elif Flight.ORsn == 'B':  # 限平衡
        ST.Cells(r, 23).Value = 'Balance problem'  # 写平衡限制
    else:  # 无拉货
        ST.Cells(r, 23).Value = ''  # 写空
    MCOLst = []  # MCO集装器列表
    if Flight.MPMC != '':  # 有MCOPMC
        MCOLst.append(Flight.MPMC + 'PMC')  # 添加MCOPMC到MCO集装器列表
    if Flight.MPAG != '':  # 有MCOPAG
        MCOLst.append(Flight.MPAG + 'PAG')  # 添加MCOPAG到MCO集装器列表
    if Flight.MPLA != '':  # 有MCOPLA
        MCOLst.append(Flight.MPLA + 'PLA')  # 添加MCOPLA到MCO集装器列表
    if Flight.MAKE != '':  # 有MCOAKE
        MCOLst.append(Flight.MAKE + 'AKE')  # 添加MCOAKE到MCO集装器列表
    Len = len(MCOLst)  # MCO集装器列表长度
    MCO = ''  # MCO集装器为空
    for i in range(Len):  # 遍历列表
        MCO = MCO + MCOLst[i]  # 添加MCO集装器字符串
        if i + 1 < Len:  # 不是列表最后
            MCO = MCO + '+'  # 末尾添加加号
    ST.Cells(r, 24).Value = MCO  # 写MCO集装器
    ST.Cells(r, 25).Value = Flight.PAX  # 写人数
    ST.Cells(r, 26).Value = Flight.ULoad  # 写剩余载量
    ST.Cells(r, 27).Value = Flight.ACNo  # 写机号
    ST.Cells(r, 28).Value = Flight.Load  # 写载重
    ST.Cells(r, 29).Value = Flight.TOW  # 写起飞重量
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
    if Flight.CW != '0':  # 有货
        ST.Cells(r, 7).Value = Flight.CW  # 写计费重量
    else:  # 无货
        ST.Cells(r, 7).Value = 'NIL'  # 写计费重量为NIL
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
    if Flight.ORsn == 'P':  # 限载
        ST.Cells(r, 23).Value = 'Payload restriction'  # 写载量限制
    elif Flight.ORsn == 'S':  # 限舱位
        ST.Cells(r, 23).Value = 'Lack of space'  # 写舱位限制
    elif Flight.ORsn == 'B':  # 限平衡
        ST.Cells(r, 23).Value = 'Balance problem'  # 写平衡限制
    else:  # 无拉货
        ST.Cells(r, 23).Value = ''  # 写空
    MCOLst = []  # MCO集装器列表
    if Flight.MPMC != '':  # 有MCOPMC
        MCOLst.append(Flight.MPMC + 'PMC')  # 添加MCOPMC到MCO集装器列表
    if Flight.MPAG != '':  # 有MCOPAG
        MCOLst.append(Flight.MPAG + 'PAG')  # 添加MCOPAG到MCO集装器列表
    if Flight.MPLA != '':  # 有MCOPLA
        MCOLst.append(Flight.MPLA + 'PLA')  # 添加MCOPLA到MCO集装器列表
    if Flight.MAKE != '':  # 有MCOAKE
        MCOLst.append(Flight.MAKE + 'AKE')  # 添加MCOAKE到MCO集装器列表
    Len = len(MCOLst)  # MCO集装器列表长度
    MCO = ''  # MCO集装器为空
    for i in range(Len):  # 遍历列表
        MCO = MCO + MCOLst[i]  # 添加MCO集装器字符串
        if i + 1 < Len:  # 不是列表最后
            MCO = MCO + '+'  # 末尾添加加号
    ST.Cells(r, 24).Value = MCO  # 写MCO集装器
    ST.Cells(r, 25).Value = Flight.PAX  # 写人数
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