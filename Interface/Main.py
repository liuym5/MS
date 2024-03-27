import sys
from Interface.Variable import MSApp

sys.exit(MSApp.exec_())

# pd.set_option('display.max_rows', None)  # 显示所有行
# pd.set_option('expand_frame_repr', False)  # 显示所有列
# print(df)  # 打印数据框架

# print(sys.exc_info())  # 打印异常

# from ReadPDF.CPM.Variable import CPMULDLst
# for CPMULD in CPMULDLst:
#     print(CPMULD.__dict__)

# from ReadPDF.UCM.Variable import UCMULDLst
# for UCMULD in UCMULDLst:
#     print(UCMULD.__dict__)

# for i in range(len(List)):  # 遍历列表
#     print(List[i].__dict__)  # 打印列表

# for Shpmt in MnfstLst:
#     print(Shpmt.__dict__)
#     for ShpmtULD in Shpmt.ULDLst:
#         print(ShpmtULD.__dict__)

# for DSULD in DSULDLst:
#     print(DSULD.__dict__)
#     for DSShpmt in DSULD.ShptLst:
#         print(DSShpmt.__dict__)