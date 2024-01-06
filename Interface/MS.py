# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'MS.ui'
#
# Created by: PyQt5 UI code generator 5.13.2
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MSForm(object):
    def setupUi(self, MSForm):
        MSForm.setObjectName("MSForm")
        MSForm.resize(221, 341)
        MSForm.setMinimumSize(QtCore.QSize(221, 341))
        MSForm.setMaximumSize(QtCore.QSize(221, 341))
        self.MsgLabel = QtWidgets.QLabel(MSForm)
        self.MsgLabel.setGeometry(QtCore.QRect(0, 0, 221, 37))
        self.MsgLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.MsgLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.MsgLabel.setObjectName("MsgLabel")
        self.DateDE = QtWidgets.QDateEdit(MSForm)
        self.DateDE.setGeometry(QtCore.QRect(70, 40, 91, 31))
        self.DateDE.setObjectName("DateDE")
        self.ULDMnfstBtn = QtWidgets.QPushButton(MSForm)
        self.ULDMnfstBtn.setGeometry(QtCore.QRect(10, 80, 61, 31))
        self.ULDMnfstBtn.setObjectName("ULDMnfstBtn")

        self.retranslateUi(MSForm)
        QtCore.QMetaObject.connectSlotsByName(MSForm)

    def retranslateUi(self, MSForm):
        _translate = QtCore.QCoreApplication.translate
        MSForm.setWindowTitle(_translate("MSForm", "MS神器"))
        self.MsgLabel.setText(_translate("MSForm", "未运行"))
        self.ULDMnfstBtn.setText(_translate("MSForm", "ULD舱单"))
