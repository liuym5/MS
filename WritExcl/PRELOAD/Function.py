def WritPRELOADST(Path, Date):  # 写PRELOAD页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    PRELOADWB = XL.Workbooks.Open(Path)  # 返回PRELOAD表格对象
    PRELOADST = PRELOADWB.Worksheets('PRE-LOAD')  # 返回PRELOAD页对象
    Date = Date + ' 1:10'  # 日期
    PRELOADST.Cells(4, 3).Value = Date  # 写日期
    r = 7  # 初始行号
    from ReadExcl.Mnfst.Variable import MnfstLst
    for shpmt in MnfstLst:  # 遍历舱单对象列表
        PRELOADST.Cells(r, 1).Value = shpmt.Seq  # 写序号
        PRELOADST.Cells(r, 2).Value = shpmt.Dest  # 写目的地
        PRELOADST.Cells(r, 3).Value = shpmt.SHC  # 写特殊操作代码
        PRELOADST.Cells(r, 4).Value = shpmt.AWBNo  # 写运单号
        PRELOADST.Cells(r, 5).Value = shpmt.Pcs  # 写件数
        PRELOADST.Cells(r, 6).Value = shpmt.Weight  # 写重量
        PRELOADST.Cells(r, 7).Value = shpmt.ChgWt  # 写计费重量
        PRELOADST.Cells(r, 8).Value = shpmt.Vol  # 写体积
        PRELOADST.Cells(r, 9).Value = shpmt.Dim  # 写尺寸
        r += 1  # 行号加1
    PRELOADWB.Save()  # 保存PRELOAD表格
    PRELOADWB.Close()  # 关闭PRELOAD表格对象
    XL.Quit()  # 关闭Excel
    from ReadExcl.Mnfst.Variable import MnfstLst
    MnfstLst.clear()  # 清空MnfstLst对象列表