def WritCPMULDStkST(ULDStkPath):  # 写CPM到ULDStock页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    ULDStkWB = XL.Workbooks.Open(ULDStkPath)  # 返回ULDStock表格对象
    ULDStkST = ULDStkWB.Worksheets('ULD Stock')  # 返回ULD Stock页对象
    WritCPMULD(ULDStkST, 'PMC', 3, 9)  # 写PMC
    WritCPMULD(ULDStkST, 'PAG', 11, 15)  # 写PAG
    WritCPMULD(ULDStkST, 'PLA', 15, 19)  # 写PLA
    WritCPMULD(ULDStkST, 'AKE', 19, 28)  # 写AKE
    ULDStkWB.Save()  # 保存ULDStock表格
    ULDStkWB.Close()  # 关闭ULDStock表格对象
    XL.Quit()  # 关闭Excel

def WritCPMULD(ST, Type, r1, r2):  # 写CPMULD
    ULDLst = ReadCPMULD(Type)  # 返回ULDLst
    ULDLstLen = len(ULDLst)  # ULDLst长度
    if ULDLstLen > 0:  # ULDLst长度大于0
        Count = int(ST.Cells(r1, 7).Text) + ULDLstLen  # 最新数量
        ST.Cells(r1, 7).Value = Count  # 更新数量
    else:  # ULDLst长度等于0
        return
    for r in range(r1, r2):  # 遍历行
        if len(ULDLst) == 0:  # ULDLst长度为0
            break
        for c in range(2, 7):  # 遍历列
            if len(ULDLst) == 0:  # ULDLst长度为0
                break
            No = ST.Cells(r, c).Text  # 号
            if No == '':  # 没有号
                for cpmuld in ULDLst:  # 遍历ULDLst
                    ST.Cells(r, c).NumberFormat = '@'  # 设置单元格格式为文本
                    ST.Cells(r, c).Value = cpmuld.No  # 写No
                    ST.Cells(r, c).Interior.ColorIndex = 6  # 单元格背景颜色为6黄色
                    ST.Cells(r, c).HorizontalAlignment = 3  # 单元格为3水平居中
                    if cpmuld.Type == 'AKE' and cpmuld.Content in ('B', 'X'):  # 类型为AKE并且内容为B或X
                        ST.Cells(r, c).Font.ColorIndex = 3  # 单元格文字颜色为3红色
                    else:  # 货运相关ULD
                        ST.Cells(r, c).Font.ColorIndex = 1  # 单元格文字颜色为1黑色
                    del ULDLst[0]  # 删除ULDLst第0项
                    break

def ReadCPMULD(Type):  # 返回ULDLst
    ULDLst = []  # ULDLst
    from ReadTXT.CPM.Variable import CPMULDLst
    for cpmuld in CPMULDLst:  # 遍历CPMULDLst
        if cpmuld.Type == Type and cpmuld.Owner != 'DS':  # 是对应类型并且所有人不为DS
            ULDLst.append(cpmuld)  # 添加CPMULD对象到ULDLst
    return ULDLst  # 返回ULDLst

def WritUCMULDStkST(ULDStkPath):  # 写UCM到ULDStock页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    ULDStkWB = XL.Workbooks.Open(ULDStkPath)  # 返回ULDStock表格对象
    ULDStkST = ULDStkWB.Worksheets('ULD Stock')  # 返回ULD Stock页对象
    WritUCMULD(ULDStkST, 'PMC', 3, 9)  # 写PMC
    WritUCMULD(ULDStkST, 'PAG', 11, 15)  # 写PAG
    WritUCMULD(ULDStkST, 'PLA', 15, 19)  # 写PLA
    ULDStkWB.Save()  # 保存ULDStock表格
    ULDStkWB.Close()  # 关闭ULDStock表格对象
    XL.Quit()  # 关闭Excel

def WritUCMULD(ST, Type, r1, r2):  # 写UCMULD
    ULDLst = ReadUCMULD(Type)  # 返回ULDLst
    ULDLstLen = len(ULDLst)  # ULDLst长度
    if ULDLstLen > 0:  # ULDLst长度大于0
        Count = int(ST.Cells(r1, 7).Text) + ULDLstLen  # 最新数量
        ST.Cells(r1, 7).Value = Count  # 更新数量
    else:  # ULDLst长度等于0
        return
    for r in range(r1, r2):  # 遍历行
        if len(ULDLst) == 0:  # ULDLst长度为0
            break
        for c in range(2, 7):  # 遍历列
            if len(ULDLst) == 0:  # ULDLst长度为0
                break
            No = ST.Cells(r, c).Text  # 号
            if No == '':  # 没有号
                for ucmuld in ULDLst:  # 遍历ULDLst
                    ST.Cells(r, c).NumberFormat = '@'  # 设置单元格格式为文本
                    ST.Cells(r, c).Value = ucmuld.No  # 写No
                    ST.Cells(r, c).Interior.ColorIndex = 6  # 单元格背景颜色为6黄色
                    del ULDLst[0]  # 删除ULDLst第0项
                    break

def ReadUCMULD(Type):  # 返回ULDLst
    ULDLst = []  # ULDLst
    from ReadTXT.UCM951.Variable import UCMULDLst
    for ucmuld in UCMULDLst:  # 遍历UCMULDLst
        if ucmuld.Type == Type and ucmuld.Owner != 'DS':  # 是对应类型并且所有人不为DS
            ULDLst.append(ucmuld)  # 添加UCMULD对象到ULDLst
    return ULDLst  # 返回ULDLst

def DelLWSULDStkST(ULDStkPath):  # 删除LWS集装器在ULD Stock页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    ULDStkWB = XL.Workbooks.Open(ULDStkPath)  # 返回ULDStock表格对象
    ULDStkST = ULDStkWB.Worksheets('ULD Stock')  # 返回ULD Stock页对象
    DelLWSULD(ULDStkST, 'PMC', 3, 9)  # 删PMC
    DelLWSULD(ULDStkST, 'PAG', 11, 15)  # 删PAG
    DelLWSULD(ULDStkST, 'PLA', 15, 19)  # 删PLA
    DelLWSULD(ULDStkST, 'AKE', 19, 28)  # 删AKE
    ULDStkWB.Save()  # 保存ULDStock表格
    ULDStkWB.Close()  # 关闭ULDStock表格对象
    XL.Quit()  # 关闭Excel

def DelLWSULD(ST, Type, r1, r2):  # 删除LWSULD
    ULDLst = ReadUCMULD(Type)  # 返回ULDLst
    ULDLstLen = len(ULDLst)  # ULDLst长度
    NoNoTF = False  # 无号为否
    ColorULDLst = []  # ColorULD列表
    if ULDLstLen > 0:  # ULDLst长度大于0
        Count = int(ST.Cells(r1, 7).Text) - ULDLstLen  # 最新数量
        ST.Cells(r1, 7).Value = Count  # 更新数量
    else:  # ULDLst长度等于0
        return
    for r in range(r1, r2):  # 遍历行
        if NoNoTF:  # 无号为是
            break
        for c in range(2, 7):  # 遍历列
            No = ST.Cells(r, c).Text  # 号
            if No == '':  # 号为空
                NoNoTF = True  # 无号为是
                break
            if NoTF(ULDLst, No) == False:  # 无号
                Font = ST.Cells(r, c).Font.ColorIndex  # 单元格文字颜色号
                Interior = ST.Cells(r, c).Interior.ColorIndex  # 单元格背景颜色号
                from WritExcl.ULDStock.Class import ColorULD
                ColorULDTmp = ColorULD(No, Font, Interior)  # 创建ColorULD对象
                ColorULDLst.append(ColorULDTmp)  # 添加ColorULD对象到ColorULD对象列表
                ST.Cells(r, c).Value = ''  # 写空
                ST.Cells(r, c).Interior.ColorIndex = 2  # 单元格背景颜色为2白色
            else:  # 有号
                if ST.Cells(r, c).Interior.ColorIndex == 3:  # 单元格背景颜色为3红色
                    print(Type + No + 'MS进港时板箱号错误！！！')
                ST.Cells(r, c).Value = ''  # 写空
                ST.Cells(r, c).Interior.ColorIndex = 2  # 单元格背景颜色为2白色
    NoStkPrt(ULDLst)  # 打印不在库存集装器号
    import operator
    cmpfun = operator.attrgetter('No')  # 参数为排序依据的属性，可以有多个，这里优先id，使用时按需求改换参数即可
    ColorULDLst.sort(key=cmpfun)  # 使用时改变列表名即可
    WritLeftULD(ColorULDLst, ST, r1, r2)  # 写剩余ULD

def NoTF(ULDLst, No):  # 是否有号
    i = 0  # 下标为0
    for ucmuld in ULDLst:  # 遍历ULDLst
        if ucmuld.No == No:  # 有号
            del ULDLst[i]  # 删除ULDLst第i项
            return True  # 有号
        i += 1  # 下标加1
    return False  # 无号

def NoStkPrt(ULDLst):  # 打印不在库存集装器号
    for ucmuld in ULDLst:  # 遍历ULDLst
        print(ucmuld.FullULDNo + '不在库存')  # 打印不在库存集装器号

def WritLeftULD(ColorULDLst, ST, r1, r2):  # 写剩余ULD
    ColorULDLstLen = len(ColorULDLst)  # ColorULDLst长度
    if ColorULDLstLen > 0:  # ColorULDLst长度大于0
        ST.Cells(r1, 7).Value = ColorULDLstLen  # 写数量
    else:  # ColorULDLstLen长度等于0
        return
    for r in range(r1, r2):  # 遍历行
        if len(ColorULDLst) == 0:  # ColorULDLst长度为0
            break
        for c in range(2, 7):  # 遍历列
            if len(ColorULDLst) == 0:  # ColorULDLst长度为0
                break
            for coloruld in ColorULDLst:  # 遍历ColorULDLst
                ST.Cells(r, c).NumberFormat = '@'  # 设置单元格格式为文本
                ST.Cells(r, c).Value = coloruld.No  # 写No
                ST.Cells(r, c).Font.ColorIndex = coloruld.Font  # 设置单元格文字颜色
                ST.Cells(r, c).Interior.ColorIndex = coloruld.Interior  # 设置单元格背景颜色
                del ColorULDLst[0]  # 删除ColorULDLst第0项
                break

def ChkUCMULDStkST(ULDStkPath):  # 检查UCMULD在ULD Stock页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    ULDStkWB = XL.Workbooks.Open(ULDStkPath)  # 返回ULDStock表格对象
    ULDStkST = ULDStkWB.Worksheets('ULD Stock')  # 返回ULD Stock页对象
    ChkUCMULD(ULDStkST, 'PMC', 3, 9)  # 检查PMC
    ChkUCMULD(ULDStkST, 'PAG', 11, 15)  # 检查PAG
    ChkUCMULD(ULDStkST, 'PLA', 15, 19)  # 检查PLA
    ChkUCMULD(ULDStkST, 'AKE', 19, 28)  # 检查AKE
    ULDStkWB.Save()  # 保存ULDStock表格
    ULDStkWB.Close()  # 关闭ULDStock表格对象
    XL.Quit()  # 关闭Excel

def ChkUCMULD(ST, Type, r1, r2):  # 检查UCMULD
    ULDLst = ReadUCMULD(Type)  # 返回ULDLst
    ULDLstLen = len(ULDLst)  # ULDLst长度
    if ULDLstLen > 0:  # ULDLst长度大于0
        for r in range(r1, r2):  # 遍历行
            if len(ULDLst) == 0:  # ULDLst长度为0
                break
            for c in range(2, 7):  # 遍历列
                if len(ULDLst) == 0:  # ULDLst长度为0
                    break
                Color = ST.Cells(r, c).Interior.ColorIndex  # 单元格背景颜色号
                No = ST.Cells(r, c).Text  # 号
                if Color == 6:  # 单元格背景颜色为6黄色
                    i = 0  # 下标为0
                    for ucmuld in ULDLst:  # 遍历ULDLst
                        if No == ucmuld.No:  # 有号
                            ST.Cells(r, c).Interior.ColorIndex = 4  # 单元格背景颜色为4绿色
                            del ULDLst[i]  # 删除ULDLst第i项
                            break
                        i += 1  # 下标加1
    NoStkPrt(ULDLst)  # 打印不在库存集装器号