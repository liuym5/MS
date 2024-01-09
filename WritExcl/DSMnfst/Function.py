def WritDSMnfst(MnfstPath):  # 写DS舱单页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    MnfstWB = XL.Workbooks.Open(MnfstPath)  # 返回舱单副本表格对象
    DSMnfstST = MnfstWB.Worksheets('DS舱单')  # 返回DS舱单页对象
    Ser = 0  # 得到初始当前序号
    r = 2  # 得到初始行号
    from WritExcl.DSMnfst.Variable import DSULDLst
    for DSULD in DSULDLst:  # 遍历DS集装器对象列表
        for DSShpmt in DSULD.ShptLst:  # 遍历货物对象列表
            if DSULD.Ser > Ser:  # 新序号集装器
                Ser += 1  # 当前序号加1
                DSMnfstST.Cells(r, 1).Value = DSULD.Ser  # 写序号号
                DSMnfstST.Cells(r, 2).Value = DSULD.Type  # 写类型
                DSMnfstST.Cells(r, 3).Value = DSULD.No  # 写号
                DSMnfstST.Cells(r, 4).Value = DSULD.Owner  # 写所有人
                DSMnfstST.Cells(r, 5).Value = DSULD.FwdTo  # 写运往
                DSMnfstST.Cells(r, 6).Value = DSShpmt.AWBNo  # 写运单号
                DSMnfstST.Cells(r, 7).Value = DSShpmt.Dest  # 写目的地
                if DSShpmt.Pcs < DSShpmt.TtlPcs:  # 集装器件数小于总件数
                    DSMnfstST.Cells(r, 8).NumberFormat = '@'  # 设置单元格格式为文本
                    DSMnfstST.Cells(r, 8).Value = str(DSShpmt.Pcs) + '/' + str(DSShpmt.TtlPcs)  # 写件数/总件数
                else:  # 集装器件数等于总件数
                    DSMnfstST.Cells(r, 8).NumberFormat = '@'  # 设置单元格格式为文本
                    DSMnfstST.Cells(r, 8).Value = DSShpmt.TtlPcs  # 写总件数
                DSMnfstST.Cells(r, 9).Value = DSShpmt.Weight  # 写重量
            else:  # 老序号集装器
                DSMnfstST.Cells(r, 6).Value = DSShpmt.AWBNo  # 写运单号
                DSMnfstST.Cells(r, 7).Value = DSShpmt.Dest  # 写目的地
                if DSShpmt.Pcs < DSShpmt.TtlPcs:  # 集装器件数小于总件数
                    DSMnfstST.Cells(r, 8).NumberFormat = '@'  # 设置单元格格式为文本
                    DSMnfstST.Cells(r, 8).Value = str(DSShpmt.Pcs) + '/' + str(DSShpmt.TtlPcs)  # 写件数/总件数
                else:  # 集装器件数等于总件数
                    DSMnfstST.Cells(r, 8).NumberFormat = '@'  # 设置单元格格式为文本
                    DSMnfstST.Cells(r, 8).Value = DSShpmt.TtlPcs  # 写总件数
                DSMnfstST.Cells(r, 9).Value = DSShpmt.Weight  # 写重量
            r += 1  # 行号加1
    MnfstWB.Save()  # 保存舱单副本表格
    MnfstWB.Close()  # 关闭舱单副本表格对象
    XL.Quit()  # 关闭Excel

def ReadMnfstLst():  # 读取舱单对象列表
    Ser = 0  # 得到初始序列号
    from ReadExcl.Mnfst.Variable import MnfstLst
    for Shpmt in MnfstLst:  # 遍历舱单对象列表
        for ShpmtULD in Shpmt.ULDLst:  # 遍历集装器对象列表
            if ShpmtULD.Owner == 'DS':  # 是DS板
                from WritExcl.DSMnfst.Variable import DSULDLst
                DSULDTmp = FindNo(DSULDLst, ShpmtULD.No)  # 返回该号DS集装器对象
                if DSULDTmp != False:  # 返回的是DS集装器对象
                    from WritExcl.DSMnfst.Class import DSShpmt
                    DSShptTmp = DSShpmt(Shpmt.AWBNo, Shpmt.Dest, ShpmtULD.Pcs, Shpmt.Pcs, ShpmtULD.Weight)  # 创建临时DS货物对象
                    DSULDTmp.AddShpt(DSShptTmp)  # 添加货物
                else:  # 返回False没有找到DS集装器对象相对应的号
                    Ser += 1  # 序列号加1
                    from WritExcl.DSMnfst.Class import DSShpmt
                    DSShptTmp = DSShpmt(Shpmt.AWBNo, Shpmt.Dest, ShpmtULD.Pcs, Shpmt.Pcs, ShpmtULD.Weight)  # 创建临时DS货物对象
                    from WritExcl.DSMnfst.Class import DSULD
                    DSULDTmp = DSULD(Ser, ShpmtULD.Type, ShpmtULD.No, ShpmtULD.Owner, 'CAI', DSShptTmp)  # 创建临时DS集装器对象
                    DSULDLst.append(DSULDTmp)  # 添加临时DS集装器对象到DS集装器对象列表

def FindNo(DSULDLst, No):  # 返回该号DS集装器对象
    for DSULD in DSULDLst:  # 遍历DS集装器对象列表
        if DSULD.No == No:  # 找到DS集装器对象相对应的号
            return DSULD  # 返回DS集装器对象
    return False  # 返回False没有找到DS集装器对象相对应的号