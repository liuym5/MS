def WritCPMULDStkST(ULDStkPath):  # 写CPM到ULDStock页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    ULDStockWB = XL.Workbooks.Open(ULDStkPath)  # 返回ULDStock表格对象
    ULDStockST = ULDStockWB.Worksheets('ULD Stock')  # 返回ULD Stock页对象
    WritCPMULD(ULDStockST, 'PMC', 3, 9)  # 写PMC
    WritCPMULD(ULDStockST, 'PAG', 11, 14)  # 写PAG
    WritCPMULD(ULDStockST, 'PLA', 14, 18)  # 写PLA
    WritCPMULD(ULDStockST, 'AKE', 18, 27)  # 写AKE
    ULDStockWB.Save()  # 保存ULDStock表格
    ULDStockWB.Close()  # 关闭ULDStock表格对象
    XL.Quit()  # 关闭Excel

def WritCPMULD(Sheet, Type, r1, r2):  # 写CPMULD
    ULDLst = ReadCPMULD(Type)  # 返回ULDLst
    ULDLstLen = len(ULDLst)  # ULDLst长度
    if ULDLstLen > 0:  # ULDLst长度大于0
        Count = int(Sheet.Cells(r1, 7).Text) + ULDLstLen  # 最新数量
        Sheet.Cells(r1, 7).Value = Count  # 更新数量
    else:  # ULDLst长度等于0
        return
    for r in range(r1, r2):  # 遍历行
        if len(ULDLst) == 0:  # ULDLst长度为0
            break
        for c in range(2, 7):  # 遍历列
            if len(ULDLst) == 0:  # ULDLst长度为0
                break
            No = Sheet.Cells(r, c).Text  # 号
            if No == '':  # 没有号
                for CPMULD in ULDLst:  # 遍历ULDLst
                    Sheet.Cells(r, c).Value = CPMULD.No  # 写No
                    Sheet.Cells(r, c).Interior.ColorIndex = 6  # 单元格背景颜色为6黄色
                    if CPMULD.Type == 'AKE' and CPMULD.Content in ('B', 'X'):  # 类型为AKE并且内容为B或X
                        Sheet.Cells(r, c).Font.ColorIndex = 3  # 单元格文字颜色为3红色
                    del ULDLst[0]  # 删除ULDLst第0项
                    break

def ReadCPMULD(Type):  # 返回ULDLst
    ULDLst = []  # ULDLst
    from ReadPDF.CPM.Variable import CPMULDLst
    for CPMULD in CPMULDLst:  # 遍历CPMULDLst
        if CPMULD.Type == Type and CPMULD.Owner == 'MS':  # 是对应类型并且所有人为MS
            ULDLst.append(CPMULD)  # 添加CPMULD对象到ULDLst
    return ULDLst  # 返回ULDLst

def WritUCMULDStkST(ULDStkPath):  # 写UCM到ULDStock页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    ULDStockWB = XL.Workbooks.Open(ULDStkPath)  # 返回ULDStock表格对象
    ULDStockST = ULDStockWB.Worksheets('ULD Stock')  # 返回ULD Stock页对象
    WritUCMULD(ULDStockST, 'PMC', 3, 9)  # 写PMC
    WritUCMULD(ULDStockST, 'PAG', 11, 14)  # 写PAG
    WritUCMULD(ULDStockST, 'PLA', 14, 18)  # 写PLA
    ULDStockWB.Save()  # 保存ULDStock表格
    ULDStockWB.Close()  # 关闭ULDStock表格对象
    XL.Quit()  # 关闭Excel

def WritUCMULD(Sheet, Type, r1, r2):  # 写UCMULD
    ULDLst = ReadUCMULD(Type)  # 返回ULDLst
    ULDLstLen = len(ULDLst)  # ULDLst长度
    if ULDLstLen > 0:  # ULDLst长度大于0
        Count = int(Sheet.Cells(r1, 7).Text) + ULDLstLen  # 最新数量
        Sheet.Cells(r1, 7).Value = Count  # 更新数量
    else:  # ULDLst长度等于0
        return
    for r in range(r1, r2):  # 遍历行
        if len(ULDLst) == 0:  # ULDLst长度为0
            break
        for c in range(2, 7):  # 遍历列
            if len(ULDLst) == 0:  # ULDLst长度为0
                break
            No = Sheet.Cells(r, c).Text  # 号
            if No == '':  # 没有号
                for UCMULD in ULDLst:  # 遍历ULDLst
                    Sheet.Cells(r, c).Value = UCMULD.No  # 写No
                    Sheet.Cells(r, c).Interior.ColorIndex = 6  # 单元格背景颜色为6黄色
                    del ULDLst[0]  # 删除ULDLst第0项
                    break

def ReadUCMULD(Type):  # 返回ULDLst
    ULDLst = []  # ULDLst
    from ReadPDF.UCM.Variable import UCMULDLst
    for UCMULD in UCMULDLst:  # 遍历UCMULDLst
        if UCMULD.Type == Type and UCMULD.Owner != 'DS':  # 是对应类型并且所有人不为DS
            ULDLst.append(UCMULD)  # 添加UCMULD对象到ULDLst
    return ULDLst  # 返回ULDLst

def DelLWSULDStkST(ULDStkPath):  # 删除LWS集装器在ULD Stock页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    ULDStockWB = XL.Workbooks.Open(ULDStkPath)  # 返回ULDStock表格对象
    ULDStockST = ULDStockWB.Worksheets('ULD Stock')  # 返回ULD Stock页对象
    DelLWSULD(ULDStockST, 'PMC', 3, 9)  # 删PMC
    DelLWSULD(ULDStockST, 'PAG', 11, 14)  # 删PAG
    DelLWSULD(ULDStockST, 'PLA', 14, 18)  # 删PLA
    DelLWSULD(ULDStockST, 'AKE', 18, 27)  # 删AKE
    ULDStockWB.Save()  # 保存ULDStock表格
    ULDStockWB.Close()  # 关闭ULDStock表格对象
    XL.Quit()  # 关闭Excel

def DelLWSULD(Sheet, Type, r1, r2):  # 删除LWSULD
    ULDLst = ReadUCMULD(Type)  # 返回ULDLst
    ULDLstLen = len(ULDLst)  # ULDLst长度
    No = ''  # 号
    NoLst = []  # 号列表
    ColorIndexLst = []  # 颜色号列表
    if ULDLstLen > 0:  # ULDLst长度大于0
        Count = int(Sheet.Cells(r1, 7).Text) - ULDLstLen  # 最新数量
        Sheet.Cells(r1, 7).Value = Count  # 更新数量
    else:  # ULDLst长度等于0
        return
    for r in range(r1, r2):  # 遍历行
        if No == '':  # 号为空
            break
        for c in range(2, 7):  # 遍历列
            No = Sheet.Cells(r, c).Text  # 号
            if No == '':  # 号为空
                break
            if NoTF(ULDLst, No):  # 有号
                Sheet.Cells(r, c).Value = ''  # 写空
                Sheet.Cells(r, c).Interior.ColorIndex = 2  # 单元格背景颜色为2白色
                continue
            NoLst.append(No)  # 添加号到号列表
            ColorIndex = Sheet.Cells(r, c).Interior.ColorIndex  # 单元格背景颜色号
            ColorIndexLst.append(ColorIndex)  # 添加颜色号到颜色号列表
            Sheet.Cells(r, c).Value = ''  # 写空
            Sheet.Cells(r, c).Interior.ColorIndex = 2  # 单元格背景颜色为2白色
    WritLeftULD(NoLst, ColorIndexLst, Sheet, r1, r2)  # 写剩余ULD

def NoTF(ULDLst, No):  # 是否有号
    for UCMULD in ULDLst:  # 遍历ULDLst
        if UCMULD.No == No:  # 有号
            del ULDLst[0]  # 删除ULDLst第0项
            return True  # 有号
    return False  # 无号

def WritLeftULD(NoLst, ColorIndexLst, Sheet, r1, r2):  # 写剩余ULD
    NoLstLen = len(NoLst)  # NoLst长度
    if NoLstLen > 0:  # NoLst长度大于0
        Sheet.Cells(r1, 7).Value = NoLstLen  # 写数量
    else:  # NoLst长度等于0
        return
    for r in range(r1, r2):  # 遍历行
        if len(NoLst) == 0:  # NoLst长度为0
            break
        for c in range(2, 7):  # 遍历列
            if len(NoLst) == 0:  # NoLst长度为0
                break
            for No in NoLst:  # 遍历NoLst
                Sheet.Cells(r, c).Value = No  # 写No
                del NoLst[0]  # 删除NoLst第0项
                break
            for ColorIndex in ColorIndexLst:  # 遍历ColorIndexLst
                Sheet.Cells(r, c).Interior.ColorIndex = ColorIndex  # 设置单元格背景颜色
                del ColorIndexLst[0]  # 删除ColorIndexLst第0项
                break