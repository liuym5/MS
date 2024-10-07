def ReadULDStkST(Path):  # 读取ULD Stock页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    ULDStkWB = XL.Workbooks.Open(Path)  # 返回ULDStock表格对象
    ULDStkST = ULDStkWB.Worksheets('ULD Stock')  # 返回ULD Stock页对象
    ReadStk(ULDStkST, 'PMC', 3, 9)  # 读取PMC
    ReadStk(ULDStkST, 'PAG', 11, 16)  # 读取PAG
    ReadStk(ULDStkST, 'PLA', 16, 20)  # 读取PLA
    ReadStk(ULDStkST, 'AKE', 20, 29)  # 读取AKE
    ULDStkWB.Save()  # 保存ULDStock表格
    ULDStkWB.Close()  # 关闭ULDStock表格对象
    XL.Quit()  # 关闭Excel
    return ReadSCM()  # 读取SCM,返回SCM

def ReadStk(ST, Type, r1, r2):  # 读取ULD
    NoNoTF = False  # 无号为否
    for r in range(r1, r2):  # 遍历行
        if NoNoTF:  # 无号为是
            break
        for c in range(2, 7):  # 遍历列
            No = ST.Cells(r, c).Text  # 号
            if No == '':  # 号为空
                NoNoTF = True  # 无号为是
                break
            ULDNo = No  # 号
            Owner = 'MS'  # 所有人为MS
            if No[-2:] in ('R7', 'R9', 'C6'):  # 所有人为R7或R9或C6
                ULDNo = No[:5]  # 号
                Owner = No[-2:]  # 所有人
            from ReadTXT.UCM951.Class import UCMULD
            UCMULDTmp = UCMULD(Type, ULDNo, Owner)  # 创建UCMULD对象
            from ReadTXT.UCM951.Variable import UCMULDLst
            UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表

def ReadSCM():  # 读取SCM,返回SCM
    SCM = 'SCM\n' + ReadDate()  # 读取SCM PACTL文本文件,返回日期
    SCM = SCM + ReadULD('AKE')  # 添加AKE
    SCM = SCM + ReadULD('PAG')  # 添加PAG
    SCM = SCM + ReadULD('PLA')  # 添加PLA
    SCM = SCM + ReadULD('PMC')  # 添加PMC
    return SCM  # 返回SCM

def ReadDate():  # 读取SCM PACTL文本文件,返回日期
    SCMPACTLFilePath = 'C:/Files/MS/ULD/SCM/SCM PACTL.TXT'  # SCM PACTL文件路径
    from ReadTXT.Function import ReadTXT
    SCMPACTL = ReadTXT(SCMPACTLFilePath)  # 读取TXT文件,返回文本
    i = SCMPACTL.find('PVG.')  # 找PVG.字符串
    if i > -1:  # 找到PVG.字符串
        Date = SCMPACTL[i:i + 15]  # 日期
    return Date

def ReadULD(Type):  # 读取ULD,返回ULD
    from WritExcl.ULDStk.Function import ReadUCMULD
    ULDLst = ReadUCMULD(Type)  # 返回ULDLst
    ULDLstLen = len(ULDLst)  # ULDLst长度
    if ULDLstLen > 0:  # 有该类型ULD
        ULD = '.' + Type + '.' + ULDLst[0].No + ULDLst[0].Owner  # 首ULD
        if ULDLstLen > 1:  # 有至少2个该类型ULD
            for i in range(1, ULDLstLen):  # 从第1项开始遍历ULDLst
                ULD = ULD + '/' + ULDLst[i].No + ULDLst[i].Owner  # 添加ULD
                if (i + 1) % 6 == 0:  # 个数除6余0
                    ULD = ULD + '\n'  # 添加换行符
        ULD = ULD.rstrip() + '.T' + str(ULDLstLen) + '\n'  # 去掉末尾换行符,添加数量
    return ULD

def ReadUnilodeST(Path, Date):  # 读取Unilode页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    ULDStkWB = XL.Workbooks.Open(Path)  # 返回ULDStock表格对象
    UnilodeST = ULDStkWB.Worksheets('Unilode')  # 返回Unilode页对象
    from Interface.Variable import CfgMS
    ReadUnilode(UnilodeST, 'PMC', CfgMS.PMCTup2)  # 读取PMC
    ReadUnilode(UnilodeST, 'PAG', CfgMS.PAGTup2)  # 读取PAG
    ULDStkWB.Save()  # 保存ULDStock表格
    ULDStkWB.Close()  # 关闭ULDStock表格对象
    XL.Quit()  # 关闭Excel
    return CreateUnilodeSCM(Date)  # 创建UnilodeSCM,返回SCM

def ReadUnilode(ST, Type, Tup2):  # 读取Unilode
    NoNoTF = False  # 无号为否
    for r in range(Tup2[0], Tup2[1]):  # 遍历行
        if NoNoTF:  # 无号为是
            break
        for c in range(2, 7):  # 遍历列
            No = ST.Cells(r, c).Text  # 号
            if No == '':  # 号为空
                NoNoTF = True  # 无号为是
                break
            ULDNo = No[:5]  # 号
            if Type == 'PAG' and ULDNo in ('17080', '56773'):  # PAJ的号
                continue  # 跳过PAJ的号
            Owner = No[-2:]  # 所有人
            from ReadTXT.UCM951.Class import UCMULD
            UCMULDTmp = UCMULD(Type, ULDNo, Owner)  # 创建UCMULD对象
            from ReadTXT.UCM951.Variable import UCMULDLst
            UCMULDLst.append(UCMULDTmp)  # 添加UCMULD对象到UCMULD对象列表

def CreateUnilodeSCM(Date):  # 创建Unilode SCM,返回SCM
    SCM = 'SCM\nPVG.' + Date + '/0110\n' # SCM头部信息
    SCM = SCM + ReadUnilodeULD('PMC')  # 添加PMC
    SCM = SCM + ReadUnilodeULD('PAG')  # 添加PAG
    SCM = SCM + '\n'  # 添加换行符
    return SCM  # 返回SCM

def ReadUnilodeULD(Type):  # 读取Unilode ULD,返回ULD
    from WritExcl.ULDStk.Function import ReadUCMULD
    ULDLst = ReadUCMULD(Type)  # 返回ULDLst
    ULDLstLen = len(ULDLst)  # ULDLst长度
    j = 0  # 6个集装器一行计数为0
    if ULDLstLen > 0:  # 有该类型ULD
        ULD = '.' + Type + '.' + ULDLst[0].No + ULDLst[0].Owner  # 首ULD
        j += 1  # 6个集装器一行计数加1
        if ULDLstLen > 1:  # 有至少2个该类型ULD
            for i in range(1, ULDLstLen):  # 从第1项开始遍历ULDLst
                ULD = ULD + '/' + ULDLst[i].No + ULDLst[i].Owner  # 添加ULD
                j += 1  # 6个集装器一行计数加1
                if (i + 1) % 6 == 0:  # 个数除6余0
                    ULD = ULD + '.T6\n'  # 添加.T6字符串和换行符
                    j = 0  # 6个集装器一行计数归0
        if j > 0:  # 有小于6个集装器的最后一行
            ULD = ULD + '.T' + str(j) + '\n'  # 添加有小于6个集装器的最后一行
    return ULD

