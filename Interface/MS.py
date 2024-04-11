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
        self.ULDMnfstBtn.setGeometry(QtCore.QRect(150, 80, 61, 31))
        self.ULDMnfstBtn.setObjectName("ULDMnfstBtn")
        self.DSMnfstBtn = QtWidgets.QPushButton(MSForm)
        self.DSMnfstBtn.setGeometry(QtCore.QRect(80, 120, 61, 31))
        self.DSMnfstBtn.setObjectName("DSMnfstBtn")
        self.ULD951Btn = QtWidgets.QPushButton(MSForm)
        self.ULD951Btn.setGeometry(QtCore.QRect(10, 80, 61, 31))
        self.ULD951Btn.setObjectName("ULD951Btn")
        self.LWS952Btn = QtWidgets.QPushButton(MSForm)
        self.LWS952Btn.setGeometry(QtCore.QRect(80, 80, 61, 31))
        self.LWS952Btn.setObjectName("LWS952Btn")
        self.UCM951Btn = QtWidgets.QPushButton(MSForm)
        self.UCM951Btn.setGeometry(QtCore.QRect(10, 120, 61, 31))
        self.UCM951Btn.setObjectName("UCM951Btn")

        self.retranslateUi(MSForm)
        QtCore.QMetaObject.connectSlotsByName(MSForm)

    def retranslateUi(self, MSForm):
        _translate = QtCore.QCoreApplication.translate
        MSForm.setWindowTitle(_translate("MSForm", "MS神器"))
        self.MsgLabel.setText(_translate("MSForm", "未运行"))
        self.ULDMnfstBtn.setText(_translate("MSForm", "ULD舱单"))
        self.DSMnfstBtn.setText(_translate("MSForm", "DS舱单"))
        self.ULD951Btn.setText(_translate("MSForm", "ULD951"))
        self.LWS952Btn.setText(_translate("MSForm", "LWS952"))
        self.UCM951Btn.setText(_translate("MSForm", "UCM951"))
