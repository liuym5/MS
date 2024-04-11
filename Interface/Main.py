import sys
from Interface.Variable import MSApp

sys.exit(MSApp.exec_())

# pd.set_option('display.max_rows', None)  # 显示所有行
# pd.set_option('expand_frame_repr', False)  # 显示所有列
# print(df)  # 打印数据框架

# print(sys.exc_info())  # 打印异常

# from ReadTXT.CPM.Variable import CPMULDLst
# for cpmuld in CPMULDLst:
#     print(cpmuld.__dict__)

# from ReadTXT.UCM951.Variable import UCMULDLst
# for ucmuld in UCMULDLst:
#     print(ucmuld.__dict__)

# for i in range(len(List)):  # 遍历列表
#     print(List[i].__dict__)  # 打印列表

# from ReadExcl.Mnfst.Variable import MnfstLst
# for shpmt in MnfstLst:
#     print(shpmt.__dict__)
#     for shpmtuld in shpmt.ULDLst:
#         print(shpmtuld.__dict__)

# from WritExcl.DSMnfst.Variable import DSULDLst
# for dsuld in DSULDLst:
#     print(dsuld.__dict__)
#     for dsshpmt in dsuld.ShptLst:
#         print(dsshpmt.__dict__)