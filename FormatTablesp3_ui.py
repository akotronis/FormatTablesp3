# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'FormatTablesp3.ui'
#
# Created by: PyQt4 UI code generator 4.11.4
#
# WARNING! All changes made in this file will be lost!

from PyQt4 import QtCore, QtGui

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName(_fromUtf8("Form"))
        Form.resize(398, 320)
        Form.setMinimumSize(QtCore.QSize(398, 320))
        Form.setMaximumSize(QtCore.QSize(398, 320))
        self.tblFiles = QtGui.QPushButton(Form)
        self.tblFiles.setGeometry(QtCore.QRect(210, 10, 181, 151))
        self.tblFiles.setObjectName(_fromUtf8("tblFiles"))
        self.DESC_COL = QtGui.QCheckBox(Form)
        self.DESC_COL.setGeometry(QtCore.QRect(10, 20, 171, 17))
        self.DESC_COL.setObjectName(_fromUtf8("DESC_COL"))
        self.ALTERNATE_CLR = QtGui.QCheckBox(Form)
        self.ALTERNATE_CLR.setGeometry(QtCore.QRect(10, 50, 151, 17))
        self.ALTERNATE_CLR.setChecked(True)
        self.ALTERNATE_CLR.setObjectName(_fromUtf8("ALTERNATE_CLR"))
        self.FIRST_COL_WIDTH = QtGui.QSpinBox(Form)
        self.FIRST_COL_WIDTH.setGeometry(QtCore.QRect(150, 80, 42, 22))
        self.FIRST_COL_WIDTH.setMinimum(1)
        self.FIRST_COL_WIDTH.setMaximum(999)
        self.FIRST_COL_WIDTH.setProperty("value", 30)
        self.FIRST_COL_WIDTH.setObjectName(_fromUtf8("FIRST_COL_WIDTH"))
        self.label = QtGui.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(10, 80, 131, 21))
        self.label.setObjectName(_fromUtf8("label"))
        self.label_2 = QtGui.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(10, 110, 131, 21))
        self.label_2.setObjectName(_fromUtf8("label_2"))
        self.OTHER_COL_WIDTH = QtGui.QSpinBox(Form)
        self.OTHER_COL_WIDTH.setGeometry(QtCore.QRect(150, 110, 42, 22))
        self.OTHER_COL_WIDTH.setMinimum(1)
        self.OTHER_COL_WIDTH.setMaximum(999)
        self.OTHER_COL_WIDTH.setProperty("value", 15)
        self.OTHER_COL_WIDTH.setObjectName(_fromUtf8("OTHER_COL_WIDTH"))
        self.scrollArea = QtGui.QScrollArea(Form)
        self.scrollArea.setGeometry(QtCore.QRect(10, 170, 381, 141))
        self.scrollArea.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        self.scrollArea.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setObjectName(_fromUtf8("scrollArea"))
        self.scrollAreaWidgetContents = QtGui.QWidget()
        self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 379, 139))
        self.scrollAreaWidgetContents.setObjectName(_fromUtf8("scrollAreaWidgetContents"))
        self.gridLayout = QtGui.QGridLayout(self.scrollAreaWidgetContents)
        self.gridLayout.setContentsMargins(5, 5, 5, 4)
        self.gridLayout.setSpacing(10)
        self.gridLayout.setObjectName(_fromUtf8("gridLayout"))
        self.outputLabel = QtGui.QLabel(self.scrollAreaWidgetContents)
        self.outputLabel.setText(_fromUtf8(""))
        self.outputLabel.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignTop)
        self.outputLabel.setWordWrap(False)
        self.outputLabel.setObjectName(_fromUtf8("outputLabel"))
        self.gridLayout.addWidget(self.outputLabel, 0, 0, 1, 1)
        self.scrollArea.setWidget(self.scrollAreaWidgetContents)
        self.LINES_BETWEEN_TABLES = QtGui.QSpinBox(Form)
        self.LINES_BETWEEN_TABLES.setGeometry(QtCore.QRect(150, 140, 42, 22))
        self.LINES_BETWEEN_TABLES.setMinimum(1)
        self.LINES_BETWEEN_TABLES.setMaximum(999)
        self.LINES_BETWEEN_TABLES.setProperty("value", 1)
        self.LINES_BETWEEN_TABLES.setObjectName(_fromUtf8("LINES_BETWEEN_TABLES"))
        self.label_3 = QtGui.QLabel(Form)
        self.label_3.setGeometry(QtCore.QRect(10, 140, 131, 21))
        self.label_3.setObjectName(_fromUtf8("label_3"))

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        Form.setWindowTitle(_translate("Form", "Format Tables", None))
        self.tblFiles.setText(_translate("Form", "Choose Table Files", None))
        self.DESC_COL.setText(_translate("Form", "Print Description Column", None))
        self.ALTERNATE_CLR.setText(_translate("Form", "Alternate Row Color", None))
        self.label.setText(_translate("Form", "First Column Width", None))
        self.label_2.setText(_translate("Form", "Other Columns Width", None))
        self.label_3.setText(_translate("Form", "Lines Between Tables", None))

