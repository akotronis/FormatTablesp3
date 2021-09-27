# -*- coding: utf-8 -*-
import sys
import datetime
from PyQt4 import QtGui
from FormatTablesp3_ui import Ui_Form
from Classes import * 


######################### Expired Date ################################
exp_date = u'30/10/2099 10:00:00'
exp_date = datetime.datetime.strptime(exp_date, '%d/%m/%Y %H:%M:%S') # day: zero-padded (01), month: zero-padded (01)
                                                                     # year: with century (2019)
                                                                     # hour: (24-hour clock) zero-padded (08)
                                                                     # minute: zero-padded (01)
                                                                     # second: zero-padded (01)
now_date = datetime.datetime.now()
#######################################################################

class MyForm(QtGui.QMainWindow):
    def __init__(self, parent=None):
        QtGui.QWidget.__init__(self, parent)
        self.ui = Ui_Form()
        self.ui.setupUi(self)
        self.print_output('Choose table files to format.\nNew file will have the name "Formatted_Tables"')
        self.exported_filname = 'Formatted_Tables'

        # Connect button to get file function
        self.ui.tblFiles.clicked.connect(self.getfile)

    def print_output(self, outp):
        self.ui.outputLabel.clear()
        if not isinstance(outp, str):
            outp = str(outp)
        self.ui.outputLabel.setText(outp)
        QtGui.QApplication.processEvents()

    def getfile(self):
        ########### OPTIONS ###########
        DESC_COL = self.ui.DESC_COL.isChecked()
        ALTERNATE_CLR = self.ui.ALTERNATE_CLR.isChecked()
        FIRST_COL_WIDTH = self.ui.FIRST_COL_WIDTH.value()
        OTHER_COL_WIDTH = self.ui.OTHER_COL_WIDTH.value()
        LINES_BETWEEN_TABLES = self.ui.LINES_BETWEEN_TABLES.value()
        MIN3_COL = rgb2hex((255,186,196))
        MIN2_COL = rgb2hex((255,123,123))
        MIN1_COL = rgb2hex((255,186,196))
        PLS1_COL = rgb2hex((210,242,212))
        PLS2_COL = rgb2hex((123,227,130))
        PLS3_COL = rgb2hex((38,204,0))
        init = "C:\\"
        ###############################
        self.ui.outputLabel.clear()
        imported_filenames = QtGui.QFileDialog.getOpenFileNames(self, 'Open file', init, 'CSV files *.csv')
        imported_filenames = [fl for fl in imported_filenames]
        if not imported_filenames or len(imported_filenames) > 2:
            self.print_output('Choose AT MOST THREE files')
            return
        self.print_output('Processing...')

        mf = MakeFile(
            imported_filenames,
            name=self.exported_filname,
            DESC_COL=DESC_COL,
            ALTERNATE_CLR=ALTERNATE_CLR,
            FIRST_COL_WIDTH=FIRST_COL_WIDTH,
            OTHER_COL_WIDTH=OTHER_COL_WIDTH,
            LINES_BETWEEN_TABLES=LINES_BETWEEN_TABLES,
            MIN3_COL=MIN3_COL,
            MIN2_COL=MIN2_COL,
            MIN1_COL=MIN1_COL,
            PLS1_COL=PLS1_COL,
            PLS2_COL=PLS2_COL,
            PLS3_COL=PLS3_COL,
            output = self.print_output,
        )
        mf.make_content()




if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    myapp = MyForm()
    myapp.show()

    if now_date > exp_date:
        msg = QtGui.QMessageBox()
        msg.setIcon(QtGui.QMessageBox.Critical)
        msg.setWindowTitle(u' ')
        msg.setText(u'Trial Period Expired')
        msg.setStandardButtons(QtGui.QMessageBox.Ok)
        msg.exec_()
        sys.exit()

    sys.exit(app.exec_())