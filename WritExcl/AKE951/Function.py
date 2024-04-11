def WritAKE951ST(AKE951Path, Date):  # 写AKE951页
    import win32com.client
    XL = win32com.client.gencache.EnsureDispatch('Excel.Application')  # 调用Excel
    XL.Visible = False  # 表格不可见
    AKE951WB = XL.Workbooks.Open(AKE951Path)  # 返回AKE951表格对象
    AKE951ST = AKE951WB.Worksheets('AKE951')  # 返回AKE951页对象
    WritAKENo(AKE951ST, Date)  # 写AKENo
    AKE951WB.Save()  # 保存AKE951表格
    AKE951WB.Close()  # 关闭AKE951表格对象
    XL.Quit()  # 关闭Excel

def WritAKENo(ST, Date):  # 写AKENo
    AKECount = 0  # AKE计数初始化为0
    r = ST.UsedRange.Rows.Count  # 末行行号
    c = 1  # 初始列号为1
    from ReadTXT.CPM.Variable import CPMULDLst
    for cpmuld in CPMULDLst:  # 遍历CPMULDLst
        if cpmuld.Type == 'AKE':  # 类型为AKE
            AKECount += 1  # AKE计数加1
            if AKECount == 1:  # AKE计数为1
                r += 1  # 末尾添加1行
                ST.Cells(r, 1).Value = Date  # 写日期
                c += 1  # 列号加1
            if c > 1 and c < 7:  # 列号大于1小于7
                ST.Columns(c).HorizontalAlignment = -4108  # 设置单元格中对齐
                ST.Cells(r, c).NumberFormat = '@'  # 设置单元格格式为文本
                ST.Cells(r, c).Value = cpmuld.No  # 写AKENo
                if cpmuld.Content in ('B', 'X'):  # 内容为B或X
                    ST.Cells(r, c).Font.ColorIndex = 3  # 单元格文字颜色为3红色
                c += 1  # 列号加1
            if c == 7:  # 列号等于7
                r += 1  # 末尾添加1行
                c = 2  # 列号为2