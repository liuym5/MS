def ReadFlight(Path, Year4, Date):  # 返回Flights表格文件当年页当天记录条对象
    import pandas as pd
    df = pd.read_excel(Path, sheet_name=Year4, header=None)  # 读取Flights表格文件当年页
    df.fillna('', inplace=True)  # NaN填充为空字符串
    FlightTmp = df[df.iloc[:, 0] == Date]  # 当前航班

    from ReadExcl.Flight.Class import Flight
    CurFlight = Flight(FlightTmp.iloc[0][0],  # 日期
                       FlightTmp.iloc[0][1],  # 机型
                       FlightTmp.iloc[0][2],  # 机号
                       FlightTmp.iloc[0][3],  # 载重
                       FlightTmp.iloc[0][4],  # 起飞重量
                       FlightTmp.iloc[0][5],  # 人数
                       FlightTmp.iloc[0][6], # 重量
                       FlightTmp.iloc[0][7],  # 剩余载重
                       FlightTmp.iloc[0][8],  # 计费重量
                       FlightTmp.iloc[0][9],  # 货PMC
                       FlightTmp.iloc[0][10],  # 货PAG
                       FlightTmp.iloc[0][11], # 货PLA
                       FlightTmp.iloc[0][12],  # 货AKE
                       FlightTmp.iloc[0][13],  # 行李PMC
                       FlightTmp.iloc[0][14],  # 行李PAG
                       FlightTmp.iloc[0][15],  # 行李PLA
                       FlightTmp.iloc[0][16],  # 行李AKE
                       FlightTmp.iloc[0][17],  # MCOPMC
                       FlightTmp.iloc[0][18],  # MCOPAG
                       FlightTmp.iloc[0][19],  # MCOPLA
                       FlightTmp.iloc[0][20],  # MCOAKE
                       FlightTmp.iloc[0][21],  # MCO目的地
                       FlightTmp.iloc[0][22],  # MCO件数
                       FlightTmp.iloc[0][23],  # 空PMC
                       FlightTmp.iloc[0][24],  # 空PAG
                       FlightTmp.iloc[0][25],  # 空PLA
                       FlightTmp.iloc[0][26],  # 空AKE
                       FlightTmp.iloc[0][27],  # 拉货PMC
                       FlightTmp.iloc[0][28],  # 拉货PAG
                       FlightTmp.iloc[0][29],  # 拉货PLA
                       FlightTmp.iloc[0][30],  # 拉货AKE
                       FlightTmp.iloc[0][31],  # 拉货重量
                       FlightTmp.iloc[0][32]  # 拉货原因
                     )  # 创建Flight对象
    return CurFlight  # 返回当年页当天记录条对象

def Getr(Path, SN, c, Date):  # 返回表格文件指定页指定列日期行号
    import pandas as pd
    df = pd.read_excel(Path, sheet_name=SN, header=None)  # 读取表格文件指定页
    df.fillna('', inplace=True)  # NaN填充为空字符串
    return df[df.iloc[:, c] == Date].index[0] + 1  # 返回行号