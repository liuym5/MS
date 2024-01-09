import sys
from Interface.Variable import MSApp

sys.exit(MSApp.exec_())

# for i in range(len(List)):  # 遍历列表
#     print(List[i].__dict__)  # 打印列表

# pd.set_option('display.max_rows', None)  # 显示所有行
# pd.set_option('expand_frame_repr', False)  # 显示所有列
# print(df)  # 打印数据框架

# for Shpmt in MnfstLst:  # 遍历舱单对象列表
#     print(Shpmt.__dict__)
#     for ShpmtULD in Shpmt.ULDLst:  # 遍历集装器对象列表
#         print(ShpmtULD.__dict__)

# for DSULD in DSULDLst:  # 遍历DS集装器对象列表
#     print(DSULD.__dict__)
#     for DSShpmt in DSULD.ShptLst:  # 遍历货物对象列表
#         print(DSShpmt.__dict__)

# print(sys.exc_info())  # 打印异常