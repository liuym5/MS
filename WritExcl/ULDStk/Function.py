def WritCPMULDStkST(Path):  # 写CPM到ULDStock页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    ULDStkWB = XL.Workbooks.Open(Path)  # 返回ULDStock表格对象
    ULDStkST = ULDStkWB.Worksheets('ULD Stock')  # 返回ULD Stock页对象
    from Interface.Variable import CfgMS
    WritCPMULD(ULDStkST, 'PMC', CfgMS.PMCTup2)  # 写PMC
    WritCPMULD(ULDStkST, 'PAG', CfgMS.PAGTup2)  # 写PAG
    WritCPMULD(ULDStkST, 'PLA', CfgMS.PLATup2)  # 写PLA
    WritCPMULD(ULDStkST, 'AKE', CfgMS.AKETup2)  # 写AKE
    WritCPMULD(ULDStkST, 'PAJ', CfgMS.PAGTup2)  # 写PAJ
    ULDStkWB.Save()  # 保存ULDStock表格
    ULDStkWB.Close()  # 关闭ULDStock表格对象
    XL.Quit()  # 关闭Excel

def WritCPMULD(ST, Type, Tup2):  # 写CPMULD
    ULDLst = ReadCPMULD(Type)  # 返回ULDLst
    ULDLstLen = len(ULDLst)  # ULDLst长度
    if ULDLstLen > 0:  # ULDLst长度大于0
        Count = int(ST.Cells(Tup2[0], 7).Text) + ULDLstLen  # 最新数量
        ST.Cells(Tup2[0], 7).Value = Count  # 更新数量
    else:  # ULDLst长度等于0
        return
    for r in range(Tup2[0], Tup2[1]):  # 遍历行
        if len(ULDLst) == 0:  # ULDLst长度为0
            break
        for c in range(2, 7):  # 遍历列
            if len(ULDLst) == 0:  # ULDLst长度为0
                break
            No = ST.Cells(r, c).Text  # 号
            if No == '':  # 没有号
                for cpmuld in ULDLst:  # 遍历ULDLst
                    ST.Cells(r, c).NumberFormat = '@'  # 单元格格式为文本
                    if cpmuld.Owner == 'MS':  # 所有人是MS
                        ST.Cells(r, c).Value = cpmuld.No # 写号
                    else:  # 所有人不是MS
                        ST.Cells(r, c).Value = cpmuld.No + cpmuld.Owner  # 写号所有人
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
        if cpmuld.Type == Type and cpmuld.Owner not in ('DS', 'K8'):  # 是对应类型并且所有人不为DS,K8
            ULDLst.append(cpmuld)  # 添加CPMULD对象到ULDLst
    return ULDLst  # 返回ULDLst

def WritUCMULDStkST(Path):  # 写UCM到ULDStock页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    ULDStkWB = XL.Workbooks.Open(Path)  # 返回ULDStock表格对象
    ULDStkST = ULDStkWB.Worksheets('ULD Stock')  # 返回ULD Stock页对象
    from Interface.Variable import CfgMS
    WritUCMULD(ULDStkST, 'PMC', CfgMS.PMCTup2)  # 写PMC
    WritUCMULD(ULDStkST, 'PAG', CfgMS.PAGTup2)  # 写PAG
    WritUCMULD(ULDStkST, 'PLA', CfgMS.PLATup2)  # 写PLA
    WritCPMULD(ULDStkST, 'PAJ', CfgMS.PAGTup2)  # 写PAJ
    ULDStkWB.Save()  # 保存ULDStock表格
    ULDStkWB.Close()  # 关闭ULDStock表格对象
    XL.Quit()  # 关闭Excel

def WritUCMULD(ST, Type, Tup2):  # 写UCMULD
    ULDLst = ReadUCMULD(Type)  # 返回ULDLst
    ULDLstLen = len(ULDLst)  # ULDLst长度
    if ULDLstLen > 0:  # ULDLst长度大于0
        Count = int(ST.Cells(Tup2[0], 7).Text) + ULDLstLen  # 最新数量
        ST.Cells(Tup2[0], 7).Value = Count  # 更新数量
    else:  # ULDLst长度等于0
        return
    for r in range(Tup2[0], Tup2[1]):  # 遍历行
        if len(ULDLst) == 0:  # ULDLst长度为0
            break
        for c in range(2, 7):  # 遍历列
            if len(ULDLst) == 0:  # ULDLst长度为0
                break
            No = ST.Cells(r, c).Text  # 号
            if No == '':  # 没有号
                for ucmuld in ULDLst:  # 遍历ULDLst
                    ST.Cells(r, c).NumberFormat = '@'  # 单元格格式为文本
                    ST.Cells(r, c).Value = ucmuld.No + ucmuld.Owner  # 写号所有人
                    ST.Cells(r, c).Interior.ColorIndex = 6  # 单元格背景颜色为6黄色
                    del ULDLst[0]  # 删除ULDLst第0项
                    break

def ReadUCMULD(Type):  # 返回ULDLst
    ULDLst = []  # ULDLst
    from ReadTXT.UCM951.Variable import UCMULDLst
    for ucmuld in UCMULDLst:  # 遍历UCMULDLst
        if ucmuld.Type == Type and ucmuld.Owner not in ('K8', 'DS'):  # 是对应类型并且所有人不为K8,DS
            ULDLst.append(ucmuld)  # 添加UCMULD对象到ULDLst
    return ULDLst  # 返回ULDLst

def DelULDStkST(Path, Date):  # 删除集装器在ULD Stock页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    ULDStkWB = XL.Workbooks.Open(Path)  # 返回ULDStock表格对象
    ULDStkST = ULDStkWB.Worksheets('ULD Stock')  # 返回ULD Stock页对象
    ULDStkST.Cells(1, 1).Value = 'PVG ULD STOCK ' + Date  # 写ULD Stock页单元格日期
    UnilodeST = ULDStkWB.Worksheets('Unilode')  # 返回Unilode页对象
    UnilodeST.Cells(1, 1).Value = 'PVG Unilode ULD STOCK ' + Date  # 写Unilode页单元格日期
    from Interface.Variable import CfgMS
    DelULD(ULDStkST, UnilodeST, 'PMC', CfgMS.PMCTup2)  # 删PMC
    DelULD(ULDStkST, UnilodeST, 'PAG', CfgMS.PAGTup2)  # 删PAG
    DelULD(ULDStkST, UnilodeST, 'PLA', CfgMS.PLATup2)  # 删PLA
    DelULD(ULDStkST, UnilodeST, 'AKE', CfgMS.AKETup2)  # 删AKE
    DelULD(ULDStkST, UnilodeST, 'PAJ', CfgMS.PAGTup2)  # 删PAJ
    ULDStkWB.Save()  # 保存ULDStock表格
    ULDStkWB.Close()  # 关闭ULDStock表格对象
    XL.Quit()  # 关闭Excel
    from ReadTXT.UCM951.Variable import UCMULDLst
    UCMULDLst.clear()  # 清空UCMULDLst

def DelULD(ULDStkST, UnilodeST, Type, Tup2):  # 删除ULD
    ULDLst = ReadUCMULD(Type)  # 返回ULDLst
    NoNoTF = False  # 无号为否
    ColorULDLst = []  # ColorULD列表
    if len(ULDLst) == 0:  # ULDLst长度等于0
        return
    for r in range(Tup2[0], Tup2[1]):  # 遍历行
        if NoNoTF:  # 无号为是
            break
        for c in range(2, 7):  # 遍历列
            No = ULDStkST.Cells(r, c).Text  # 号
            if No == '':  # 号为空
                NoNoTF = True  # 无号为是
                break
            ULDNo = No
            Owner = 'MS'  # 所有人为MS
            if No[-2:] in ('R7', 'R9', 'C6'):  # 所有人为R7或R9或C6
                ULDNo = No[:5]  # 号
                Owner = No[-2:]  # 所有人
            if NoTF(ULDLst, ULDNo) == False:  # 无号
                Font = ULDStkST.Cells(r, c).Font.ColorIndex  # 单元格文字颜色号
                Interior = ULDStkST.Cells(r, c).Interior.ColorIndex  # 单元格背景颜色号
                from WritExcl.ULDStk.Class import ColorULD
                ColorULDTmp = ColorULD(ULDNo, Owner, Font, Interior)  # 创建ColorULD对象
                ColorULDLst.append(ColorULDTmp)  # 添加ColorULD对象到ColorULD对象列表
                ULDStkST.Cells(r, c).Value = ''  # ULDStock页单元格写空
                ULDStkST.Cells(r, c).Interior.ColorIndex = 2  # ULDStock页单元格背景颜色为2白色
            else:  # 有号
                ColorIndex = ULDStkST.Cells(r, c).Interior.ColorIndex  # 单元格背景颜色号
                if ColorIndex in (3, 8, 33):  # 单元格背景颜色为3红色或8蓝色或33蓝色
                    print(Type + ULDNo + Owner + '进港时板箱号错误！！！')
                elif ColorIndex == 47:  # 单元格背景颜色为47紫色
                    print(Type + ULDNo + Owner + '无进港记录！！！')
                ULDStkST.Cells(r, c).Value = ''  # ULDStock页单元格写空
                ULDStkST.Cells(r, c).Interior.ColorIndex = 2  # ULDStock页单元格背景颜色为2白色
            if Type in ('PMC' 'PAG' 'PAJ'):  # 类型为PMC或PAG或PAJ
                UnilodeST.Cells(r, c).Value = ''  # Unilode页单元格写空
    NoStkPrt(ULDLst, Type)  # 打印不在库存集装器号
    import operator
    cmpfun = operator.attrgetter('No')  # 参数为排序依据的属性，可以有多个，这里优先id，使用时按需求改换参数即可
    ColorULDLst.sort(key=cmpfun)  # 使用时改变列表名即可
    WritLeftULD(ColorULDLst, ULDStkST, UnilodeST, Tup2)  # 写剩余ULD

def NoTF(ULDLst, No):  # 是否有号
    i = 0  # 下标为0
    for ucmuld in ULDLst:  # 遍历ULDLst
        if ucmuld.No == No:  # 有号
            del ULDLst[i]  # 删除ULDLst第i项
            return True  # 有号
        i += 1  # 下标加1
    return False  # 无号

def NoStkPrt(ULDLst, Type):  # 打印不在库存集装器号
    for ucmuld in ULDLst:  # 遍历ULDLst
        print(ucmuld.FullULDNo + '不在库存！！！')  # 打印不在库存集装器号
    print('总共' + str(len(ULDLst)) + '个' + Type + '不在库存！！！')

def WritLeftULD(ColorULDLst, ULDStkST, UnilodeST, Tup2):  # 写剩余ULD
    UnilodeLst = []  # Unilode列表
    ColorULDLstLen = len(ColorULDLst)  # ColorULDLst长度
    if ColorULDLstLen > 0:  # ColorULDLst长度大于0
        ULDStkST.Cells(Tup2[0], 7).Value = ColorULDLstLen  # 写ULDStock页单元格数量
    else:  # ColorULDLstLen长度等于0
        ULDStkST.Cells(Tup2[0], 7).Value = 0  # 写ULDStock页单元格数量为0
        return
    for r in range(Tup2[0], Tup2[1]):  # 遍历行
        if len(ColorULDLst) == 0:  # ColorULDLst长度为0
            break
        for c in range(2, 7):  # 遍历列
            if len(ColorULDLst) == 0:  # ColorULDLst长度为0
                break
            for coloruld in ColorULDLst:  # 遍历ColorULDLst
                ULDStkST.Cells(r, c).NumberFormat = '@'  # 单元格格式为文本
                if coloruld.Owner == 'MS':  # 所有人为MS
                    No = coloruld.No  # 号
                else:  # 所有人不为MS
                    No = coloruld.No + coloruld.Owner  # 号加所有人
                    UnilodeLst.append(No)  # 添加No到Unilode列表
                ULDStkST.Cells(r, c).Value = No  # ULDStock页单元格写号
                ULDStkST.Cells(r, c).Font.ColorIndex = coloruld.Font  # ULDStock页单元格文字颜色
                ULDStkST.Cells(r, c).Interior.ColorIndex = coloruld.Interior  # ULDStock页单元格背景颜色
                ULDStkST.Cells(r, c).HorizontalAlignment = 3  # ULDStock页单元格为3水平居中
                del ColorULDLst[0]  # 删除ColorULDLst第0项
                break
    UnilodeLstLen = len(UnilodeLst)  # UnilodeLst长度
    if UnilodeLstLen > 0:  # UnilodeLstLen长度大于0
        UnilodeST.Cells(Tup2[0], 7).Value = UnilodeLstLen  # 写Unilode页单元格数量
    else:  # UnilodeLstLen长度等于0
        UnilodeST.Cells(Tup2[0], 7).Value = 0  # 写Unilode页单元格数量为0
        return
    for r in range(Tup2[0], Tup2[1]):  # 遍历行
        if len(UnilodeLst) == 0:  # UnilodeLst长度为0
            break
        for c in range(2, 7):  # 遍历列
            if len(UnilodeLst) == 0:  # UnilodeLst长度为0
                break
            for No in UnilodeLst:  # 遍历UnilodeLst
                UnilodeST.Cells(r, c).NumberFormat = '@'  # Unilode页单元格格式为文本
                UnilodeST.Cells(r, c).Value = No  # Unilode页单元格写号
                UnilodeST.Cells(r, c).HorizontalAlignment = 3  # Unilode页单元格为3水平居中
                del UnilodeLst[0]  # 删除UnilodeLst第0项
                break

def ChkUCMULDStkST(Path):  # 检查UCMULD在ULD Stock页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    ULDStkWB = XL.Workbooks.Open(Path)  # 返回ULDStock表格对象
    ULDStkST = ULDStkWB.Worksheets('ULD Stock')  # 返回ULD Stock页对象
    from Interface.Variable import CfgMS
    ChkUCMULD(ULDStkST, 'PMC', CfgMS.PMCTup2)  # 检查PMC
    ChkUCMULD(ULDStkST, 'PAG', CfgMS.PAGTup2)  # 检查PAG
    ChkUCMULD(ULDStkST, 'PAJ', CfgMS.PAGTup2)  # 检查PAJ
    ChkUCMULD(ULDStkST, 'PLA', CfgMS.PLATup2)  # 检查PLA
    ChkUCMULD(ULDStkST, 'AKE', CfgMS.AKETup2)  # 检查AKE
    ULDStkWB.Save()  # 保存ULDStock表格
    ULDStkWB.Close()  # 关闭ULDStock表格对象
    XL.Quit()  # 关闭Excel

def ChkUCMULD(ST, Type, Tup2):  # 检查UCMULD
    ULDLst = ReadUCMULD(Type)  # 返回ULDLst
    ULDLstLen = len(ULDLst)  # ULDLst长度
    if ULDLstLen > 0:  # ULDLst长度大于0
        for r in range(Tup2[0], Tup2[1]):  # 遍历行
            if len(ULDLst) == 0:  # ULDLst长度为0
                break
            for c in range(2, 7):  # 遍历列
                if len(ULDLst) == 0:  # ULDLst长度为0
                    break
                Color = ST.Cells(r, c).Interior.ColorIndex  # 单元格背景颜色号
                No = ST.Cells(r, c).Text  # 号
                if No[-2:] in ('R7', 'R9', 'C6'):  # 所有人为R7或R9或C6
                    No = No[:5]  # 号
                if Color == 6:  # 单元格背景颜色为6黄色
                    i = 0  # 下标为0
                    for ucmuld in ULDLst:  # 遍历ULDLst
                        if No == ucmuld.No:  # 有号
                            ST.Cells(r, c).Font.ColorIndex = 1  # 单元格文字颜色为1黑色
                            ST.Cells(r, c).Interior.ColorIndex = 4  # 单元格背景颜色为4绿色
                            del ULDLst[i]  # 删除ULDLst第i项
                            break
                        i += 1  # 下标加1
    AddUCMULD(ULDLst, ST, Tup2, Type)  # 添加UCMULD

def AddUCMULD(ULDLst, ST, Tup2, Type):  # 添加UCMULD
    ULDLstLen = len(ULDLst)  # ULDLst长度
    if ULDLstLen > 0:  # ULDLst长度大于0
        Count = int(ST.Cells(Tup2[0], 7).Text) + ULDLstLen  # 最新数量
        ST.Cells(Tup2[0], 7).Value = Count  # 更新数量
        for r in range(Tup2[0], Tup2[1]):  # 遍历行
            if len(ULDLst) == 0:  # ULDLst长度为0
                break
            for c in range(2, 7):  # 遍历列
                if len(ULDLst) == 0:  # ULDLst长度为0
                    break
                No = ST.Cells(r, c).Text  # 号
                if No == '':  # 无号
                    for ucmuld in ULDLst:  # 遍历ULDLst
                        No = ucmuld.No  # 号
                        if ucmuld.Owner in ('R7', 'R9', 'C6'):  # R7R9C6板
                            No = ucmuld.No + ucmuld.Owner  # 号加所有人
                        ST.Cells(r, c).NumberFormat = '@'  # 单元格格式为文本
                        ST.Cells(r, c).Value = No  # 写号
                        ST.Cells(r, c).Interior.ColorIndex = 6  # 单元格背景颜色为6黄色
                        ST.Cells(r, c).HorizontalAlignment = 3  # 单元格为3水平居中
                        print(ucmuld.FullULDNo + '不在库存！！！')  # 打印不在库存集装器号
                        del ULDLst[0]  # 删除ULDLst首项
                        break
    print('总共' + str(ULDLstLen) + '个' + Type + '不在库存！！！')

def ChkSCMULDStkST(Path):  # 检查SCM在ULD Stock页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    ULDStkWB = XL.Workbooks.Open(Path)  # 返回ULDStock表格对象
    ULDStkST = ULDStkWB.Worksheets('ULD Stock')  # 返回ULD Stock页对象
    from Interface.Variable import CfgMS
    ChkSCM(ULDStkST, 'PMC', CfgMS.PMCTup2)  # 检查PMC
    ChkSCM(ULDStkST, 'PAG', CfgMS.PAGTup2)  # 检查PAG
    ChkSCM(ULDStkST, 'PAJ', CfgMS.PAGTup2)  # 检查PAJ
    ChkSCM(ULDStkST, 'PLA', CfgMS.PLATup2)  # 检查PLA
    ChkSCM(ULDStkST, 'AKE', CfgMS.AKETup2)  # 检查AKE
    ULDStkWB.Save()  # 保存ULDStock表格
    ULDStkWB.Close()  # 关闭ULDStock表格对象
    XL.Quit()  # 关闭Excel

def ChkSCM(ST, Type, Tup2):  # 检查SCM
    ULDLst = ReadUCMULD(Type)  # 返回ULDLst
    ULDLstLen = len(ULDLst)  # ULDLst长度
    NoNoTF = False  # 无号为否
    if ULDLstLen > 0:  # ULDLst长度大于0
        for r in range(Tup2[0], Tup2[1]):  # 遍历行
            if NoNoTF:  # 无号为是
                break
            for c in range(2, 7):  # 遍历列
                No = ST.Cells(r, c).Text  # 号
                if No == '':  # 号为空
                    NoNoTF = True  # 无号为是
                    break
                if No[-2:] in ('R7', 'R9', 'C6'):  # 所有人为R7或R9或C6
                    No = No[:5]  # 号
                i = 0  # 下标为0
                for ucmuld in ULDLst:  # 遍历ULDLst
                    if No == ucmuld.No:  # 有号
                        Color = ST.Cells(r, c).Interior.ColorIndex  # 单元格背景颜色号
                        if Color != 3:  # 单元格背景颜色不为3红色
                            ST.Cells(r, c).Interior.ColorIndex = 4  # 单元格背景颜色为4绿色
                        del ULDLst[i]  # 删除ULDLst第i项
                        break
                    i += 1  # 下标加1
                    if i == len(ULDLst):  # 无号
                        ST.Cells(r, c).Interior.ColorIndex = 8  # 单元格背景颜色为8淡蓝色
    NoStkPrt(ULDLst)  # 打印不在库存集装器号