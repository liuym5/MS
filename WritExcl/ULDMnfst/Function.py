def WritULDMnfstST(Path):  # 写集装器舱单信息页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    ULDMnfstWB = XL.Workbooks.Open(Path)  # 返回集装器舱单表格对象
    ULDMnfstST = ULDMnfstWB.Worksheets('舱单信息')  # 返回集装器舱单信息页对象
    r = 8  # 得到初始行号
    from ReadExcl.Mnfst.Variable import MnfstLst
    for shpmt in MnfstLst:  # 遍历舱单对象列表
        for shpmtuld in shpmt.ULDLst:  # 遍历集装器对象列表
            ULDMnfstST.Cells(r, 1).Value = shpmtuld.ULDNo  # 写集装器号
            ULDMnfstST.Cells(r, 2).Value = shpmtuld.Type  # 写集装器类型
            ULDMnfstST.Cells(r, 3).Value = shpmt.AWBNo  # 写运单号
            ULDMnfstST.Cells(r, 4).Value = shpmt.SHC  # 写特殊操作代码
            ULDMnfstST.Cells(r, 5).Value = shpmt.ManDesc  # 写品名
            ULDMnfstST.Cells(r, 6).Value = shpmt.Pcs  # 写总件数
            ULDMnfstST.Cells(r, 7).Value = shpmtuld.Pcs  # 写集装器件数
            ULDMnfstST.Cells(r, 8).Value = shpmtuld.Weight  # 写集装器重量
            ULDMnfstST.Cells(r, 9).Value = shpmtuld.Vol  # 写集装器体积
            ULDMnfstST.Cells(r, 10).Value = shpmtuld.ChgWt  # 写集装器计费重量
            r += 1  # 行号加1
    ULDMnfstWB.Save()  # 保存集装器舱单表格
    ULDMnfstWB.Close()  # 关闭集装器舱单表格对象
    XL.Quit()  # 关闭Excel
    MnfstLst.clear()  # 清空航班舱单Shpmt对象列表
