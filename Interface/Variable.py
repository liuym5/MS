from Config.Class import Cfg
CfgMS = Cfg()  # 配置文件类对象
from PyQt5.QtWidgets import QApplication
import sys
MSApp = QApplication(sys.argv)
from Interface.Class import MSMainForm
MSWin = MSMainForm()
MSWin.show()