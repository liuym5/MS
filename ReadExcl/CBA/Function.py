def ReadCBAST(Path):  # 读取CBA表格CBA页
    import pandas as pd
    df = pd.read_excel(Path, sheet_name='CBA')  # 读取CBA表格CBA页
    for r in range(len(df)):  # 遍历所有行
        AWBNo = df.iloc[r][1]  # 运单号
        Dim = df.iloc[r][8]  # 尺寸
        from ReadExcl.CBA.Class import CBADim
        CBADimTmp = CBADim(AWBNo, Dim)  # 创建CBADim对象
        from ReadExcl.CBA.Variable import CBADimLst
        CBADimLst.append(CBADimTmp)  # 添加到CBADim对象列表

def ReadDim():  # 读取Dim
    from ReadExcl.Mnfst.Variable import MnfstLst
    for shpmt in MnfstLst:  # 遍历MnfstLst
        i = 0  # 下标为0
        from ReadExcl.CBA.Variable import CBADimLst
        for cbadim in CBADimLst:  # 遍历CBADim
            if shpmt.AWBNo == cbadim.AWBNo:  # 找到运单号
                shpmt.Dim = cbadim.Dim  # 尺寸
                del CBADimLst[i]  # 删除CBADimLst第i项
                break
            i += 1  # 下标加1