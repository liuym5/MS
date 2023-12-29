from PyQt5.QtWidgets import QApplication
import sys
MSApp = QApplication(sys.argv)
from Interface.Class import MSMainForm
MSWin = MSMainForm()
MSWin.show()