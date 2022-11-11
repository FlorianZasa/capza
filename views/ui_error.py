# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '\\mac\Home\Desktop\myBots\capza-app\capza\views\error.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(795, 511)
        Dialog.setStyleSheet("background-color: #F7F7F7")
        self.horizontalLayout = QtWidgets.QHBoxLayout(Dialog)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setSpacing(0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setContentsMargins(10, 10, 0, 10)
        self.verticalLayout.setSpacing(20)
        self.verticalLayout.setObjectName("verticalLayout")
        self.error_sign = QtWidgets.QLabel(Dialog)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.error_sign.sizePolicy().hasHeightForWidth())
        self.error_sign.setSizePolicy(sizePolicy)
        self.error_sign.setMinimumSize(QtCore.QSize(40, 40))
        self.error_sign.setMaximumSize(QtCore.QSize(40, 40))
        font = QtGui.QFont()
        font.setPointSize(40)
        self.error_sign.setFont(font)
        self.error_sign.setText("")
        self.error_sign.setPixmap(QtGui.QPixmap("\\\\mac\\Home\\Desktop\\myBots\\capza-app\\capza\\views\\../assets/error_icon.png"))
        self.error_sign.setScaledContents(True)
        self.error_sign.setObjectName("error_sign")
        self.verticalLayout.addWidget(self.error_sign)
        self.error_msg_frame = QtWidgets.QFrame(Dialog)
        self.error_msg_frame.setStyleSheet("background: #FCFCFC;\n"
"border-radius: 10px")
        self.error_msg_frame.setObjectName("error_msg_frame")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.error_msg_frame)
        self.verticalLayout_3.setContentsMargins(10, 10, 10, 10)
        self.verticalLayout_3.setSpacing(20)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.error_lbl = QtWidgets.QLabel(self.error_msg_frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.error_lbl.sizePolicy().hasHeightForWidth())
        self.error_lbl.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Leelawadee UI")
        font.setPointSize(11)
        self.error_lbl.setFont(font)
        self.error_lbl.setStyleSheet("")
        self.error_lbl.setWordWrap(True)
        self.error_lbl.setTextInteractionFlags(QtCore.Qt.LinksAccessibleByMouse|QtCore.Qt.TextSelectableByKeyboard|QtCore.Qt.TextSelectableByMouse)
        self.error_lbl.setObjectName("error_lbl")
        self.verticalLayout_3.addWidget(self.error_lbl, 0, QtCore.Qt.AlignTop)
        self.verticalLayout.addWidget(self.error_msg_frame)
        self.horizontalLayout.addLayout(self.verticalLayout)
        self.verticalFrame_2 = QtWidgets.QFrame(Dialog)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.verticalFrame_2.sizePolicy().hasHeightForWidth())
        self.verticalFrame_2.setSizePolicy(sizePolicy)
        self.verticalFrame_2.setLayoutDirection(QtCore.Qt.RightToLeft)
        self.verticalFrame_2.setStyleSheet("background-color: #0C5C51;")
        self.verticalFrame_2.setObjectName("verticalFrame_2")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.verticalFrame_2)
        self.verticalLayout_2.setContentsMargins(10, 10, 10, 10)
        self.verticalLayout_2.setSpacing(20)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_2.addItem(spacerItem)
        self.close_error_info_btn = QtWidgets.QPushButton(self.verticalFrame_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.close_error_info_btn.sizePolicy().hasHeightForWidth())
        self.close_error_info_btn.setSizePolicy(sizePolicy)
        self.close_error_info_btn.setMinimumSize(QtCore.QSize(60, 60))
        font = QtGui.QFont()
        font.setFamily("Leelawadee UI")
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.close_error_info_btn.setFont(font)
        self.close_error_info_btn.setStyleSheet("background-color: #AEDC21;\n"
"border-radius: 10px;\n"
"text-align:center;")
        self.close_error_info_btn.setObjectName("close_error_info_btn")
        self.verticalLayout_2.addWidget(self.close_error_info_btn)
        self.horizontalLayout.addWidget(self.verticalFrame_2)

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.error_lbl.setText(_translate("Dialog", "TextLabel"))
        self.close_error_info_btn.setText(_translate("Dialog", "Ok"))
