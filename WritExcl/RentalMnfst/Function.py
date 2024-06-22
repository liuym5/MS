def WritRentalMnfstST(Path):  # 写DS舱单页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    MnfstWB = XL.Workbooks.Open(Path)  # 返回舱单副本表格对象
    DSMnfstST = MnfstWB.Worksheets('租板舱单')  # 返回DS舱单页对象
    Ser = 0  # 得到初始当前序号
    r = 2  # 得到初始行号
    from WritExcl.RentalMnfst.Variable import DSULDLst
    for dsuld in DSULDLst:  # 遍历DS集装器对象列表
        for dsshpmt in dsuld.ShptLst:  # 遍历货物对象列表
            if dsuld.Ser > Ser:  # 新序号集装器
                Ser += 1  # 当前序号加1
                DSMnfstST.Cells(r, 1).Value = dsuld.Ser  # 写序号号
                DSMnfstST.Cells(r, 2).Value = dsuld.Type  # 写类型
                DSMnfstST.Cells(r, 3).NumberFormat = '@'  # 设置单元格格式为文本
                DSMnfstST.Cells(r, 3).Value = dsuld.No  # 写号
                DSMnfstST.Cells(r, 4).Value = dsuld.Owner  # 写所有人
                DSMnfstST.Cells(r, 5).Value = dsuld.FwdTo  # 写运往
                DSMnfstST.Cells(r, 6).Value = dsshpmt.AWBNo  # 写运单号
                DSMnfstST.Cells(r, 7).Value = dsshpmt.Dest  # 写目的地
                if dsshpmt.Pcs < dsshpmt.TtlPcs:  # 集装器件数小于总件数
                    DSMnfstST.Cells(r, 8).NumberFormat = '@'  # 设置单元格格式为文本
                    DSMnfstST.Cells(r, 8).Value = str(dsshpmt.Pcs) + '/' + str(dsshpmt.TtlPcs)  # 写件数/总件数
                else:  # 集装器件数等于总件数
                    DSMnfstST.Cells(r, 8).Value = dsshpmt.TtlPcs  # 写总件数
                DSMnfstST.Cells(r, 9).Value = dsshpmt.Weight  # 写重量
            else:  # 老序号集装器
                DSMnfstST.Cells(r, 6).Value = dsshpmt.AWBNo  # 写运单号
                DSMnfstST.Cells(r, 7).Value = dsshpmt.Dest  # 写目的地
                if dsshpmt.Pcs < dsshpmt.TtlPcs:  # 集装器件数小于总件数
                    DSMnfstST.Cells(r, 8).NumberFormat = '@'  # 设置单元格格式为文本
                    DSMnfstST.Cells(r, 8).Value = str(dsshpmt.Pcs) + '/' + str(dsshpmt.TtlPcs)  # 写件数/总件数
                else:  # 集装器件数等于总件数
                    DSMnfstST.Cells(r, 8).Value = dsshpmt.TtlPcs  # 写总件数
                DSMnfstST.Cells(r, 9).Value = dsshpmt.Weight  # 写重量
            r += 1  # 行号加1
    MnfstWB.Save()  # 保存舱单副本表格
    MnfstWB.Close()  # 关闭舱单副本表格对象
    XL.Quit()  # 关闭Excel
    from ReadExcl.Mnfst.Variable import MnfstLst
    MnfstLst.clear()  # 清空航班舱单Shpmt对象列表

def ReadMnfstLst():  # 读取舱单对象列表
    Ser = 0  # 得到初始序列号
    from ReadExcl.Mnfst.Variable import MnfstLst
    for shpmt in MnfstLst:  # 遍历舱单对象列表
        for shpmtuld in shpmt.ULDLst:  # 遍历集装器对象列表
            if shpmtuld.Owner in ('K8', 'DS'):  # 是K8,DS板
                from WritExcl.RentalMnfst.Variable import DSULDLst
                DSULDTmp = FindNo(DSULDLst, shpmtuld.No)  # 返回该号DSULD对象
                if DSULDTmp != False:  # 返回的是DS集装器对象
                    from WritExcl.RentalMnfst.Class import DSShpmt
                    DSShptTmp = DSShpmt(shpmt.AWBNo, shpmt.Dest, shpmtuld.Pcs, shpmt.Pcs, shpmtuld.Weight)  # 创建DSShpmt对象
                    DSULDTmp.AddShpt(DSShptTmp)  # 添加货物
                else:  # 返回False没有找到DSULD对象相对应的号
                    Ser += 1  # 序列号加1
                    from WritExcl.RentalMnfst.Class import DSShpmt
                    DSShptTmp = DSShpmt(shpmt.AWBNo, shpmt.Dest, shpmtuld.Pcs, shpmt.Pcs, shpmtuld.Weight)  # 创建DSShpmt对象
                    from WritExcl.RentalMnfst.Class import DSULD
                    DSULDTmp = DSULD(Ser, shpmtuld.Type, shpmtuld.No, shpmtuld.Owner, 'CAI', DSShptTmp)  # 创建DSULD对象
                    DSULDLst.append(DSULDTmp)  # 添加DSULD对象到DSULD对象列表

def FindNo(DSULDLst, No):  # 返回该号DSULD对象
    for dsuld in DSULDLst:  # 遍历DSULD对象列表
        if dsuld.No == No:  # 找到DSULD对象相对应的号
            return dsuld  # 返回DSULD对象
    return False  # 返回False没有找到DSULD对象相对应的号