# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'ui/03.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets
from lanDialog import LanDialog
from aboutDialog import AboutDialog
from model import Model
import sys
import os

# ASSET PATH
# distribution
assetPath = os.path.dirname(os.path.realpath(sys.executable)) + '/assets/'

# dev 
#assetPath = 'src\\../assets/'


class Ui_MainWindow(object):
    def __init__( self ):
        # Initialize the super class
        super().__init__()
        self.model = Model(self)

        # About dialog initialization
        self.aboutDialog = QtWidgets.QDialog()
        self.ui2 = AboutDialog()
        self.ui2.setupUi(self.aboutDialog, assetPath)

        # Language select dialog initialization
        self.lanDialog = QtWidgets.QDialog()
        # IMPORTANT: ui3 have to be add self. Otherwise ui3 object will be recycled after __init__() 
        # Then the radio button can't response to the click
        self.ui3 = LanDialog(self, self.ui2)                 
        self.ui3.setupUi(self.lanDialog, assetPath)
        

    def setupUi(self, MainWindow):
        self.MainWindow = MainWindow
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(524, 394)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(assetPath + 'icon.png'), QtGui.QIcon.Normal, QtGui.QIcon.Off) # window icon
        MainWindow.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.inputLayout = QtWidgets.QVBoxLayout()
        self.inputLayout.setSizeConstraint(QtWidgets.QLayout.SetMaximumSize)
        self.inputLayout.setContentsMargins(-1, 10, -1, 20)
        self.inputLayout.setSpacing(20)
        self.inputLayout.setObjectName("inputLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setSpacing(10)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.sourceLabel = QtWidgets.QLabel(self.centralwidget)
        self.sourceLabel.setObjectName("sourceLabel")
        self.horizontalLayout.addWidget(self.sourceLabel)
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout.addWidget(self.lineEdit)

        self.browseButton1 = QtWidgets.QPushButton(self.centralwidget)
        self.browseButton1.setObjectName("browseButton1")
        self.browseButton1.clicked.connect(lambda: self.browseButtonClicked(1))   # bind click function
        
        self.horizontalLayout.addWidget(self.browseButton1)
        spacerItem = QtWidgets.QSpacerItem(30, 20, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem)
        self.sheetSelect1 = QtWidgets.QComboBox(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.sheetSelect1.sizePolicy().hasHeightForWidth())
        self.sheetSelect1.setSizePolicy(sizePolicy)
        self.sheetSelect1.setMinimumSize(QtCore.QSize(90, 0))
        self.sheetSelect1.setObjectName("sheetSelect1")
        self.horizontalLayout.addWidget(self.sheetSelect1)
        self.inputLayout.addLayout(self.horizontalLayout)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setSpacing(10)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.targetLabel = QtWidgets.QLabel(self.centralwidget)
        self.targetLabel.setObjectName("targetLabel")
        self.horizontalLayout_2.addWidget(self.targetLabel)
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.horizontalLayout_2.addWidget(self.lineEdit_2)

        self.browseButton2 = QtWidgets.QPushButton(self.centralwidget)
        self.browseButton2.setObjectName("browseButton2")
        self.browseButton2.clicked.connect(lambda: self.browseButtonClicked(2))   # bind click function

        self.horizontalLayout_2.addWidget(self.browseButton2)
        spacerItem1 = QtWidgets.QSpacerItem(30, 20, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem1)
        self.sheetSelect2 = QtWidgets.QComboBox(self.centralwidget)
        self.sheetSelect2.setMinimumSize(QtCore.QSize(90, 0))
        self.sheetSelect2.setObjectName("sheetSelect2")
        self.horizontalLayout_2.addWidget(self.sheetSelect2)
        self.inputLayout.addLayout(self.horizontalLayout_2)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setSpacing(10)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.headerNameLabel = QtWidgets.QLabel(self.centralwidget)
        self.headerNameLabel.setObjectName("headerNameLabel")
        self.horizontalLayout_3.addWidget(self.headerNameLabel)
        self.lineEdit_3 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.horizontalLayout_3.addWidget(self.lineEdit_3)
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_3.addItem(spacerItem2)

        self.startButton = QtWidgets.QPushButton(self.centralwidget)
        self.startButton.setObjectName("startButton")
        self.startButton.clicked.connect(self.start)   # bind click function

        self.horizontalLayout_3.addWidget(self.startButton)
        self.inputLayout.addLayout(self.horizontalLayout_3)
        self.verticalLayout_4.addLayout(self.inputLayout)
        self.logLabel = QtWidgets.QLabel(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(11)
        self.logLabel.setFont(font)
        self.logLabel.setTextFormat(QtCore.Qt.AutoText)
        self.logLabel.setObjectName("logLabel")
        self.verticalLayout_4.addWidget(self.logLabel)

        self.logBrowser = QtWidgets.QTextBrowser(self.centralwidget)
        self.logBrowser.setObjectName("logBrowser")

        self.verticalLayout_4.addWidget(self.logBrowser)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 524, 23))
        self.menubar.setObjectName("menubar")
        self.menuPreference = QtWidgets.QMenu(self.menubar)
        self.menuPreference.setObjectName("menuPreference")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionLanguage = QtWidgets.QAction(MainWindow)
        self.actionLanguage.setObjectName("actionLanguage")
        self.actionLanguage.triggered.connect(self.openLanDialog)    # bind trigger
        self.actionClearLog = QtWidgets.QAction(MainWindow)
        self.actionClearLog.setObjectName("actionClearLog")
        self.actionClearLog.triggered.connect(self.clearLog)    # bind trigger
        self.actionAbout = QtWidgets.QAction(MainWindow)
        self.actionAbout.setObjectName("actionAbout")
        self.actionAbout.triggered.connect(self.openAboutDialog)    # bind trigger
        self.menuPreference.addAction(self.actionLanguage)
        self.menuPreference.addAction(self.actionClearLog)
        self.menuPreference.addSeparator()
        self.menuPreference.addAction(self.actionAbout)
        self.menubar.addAction(self.menuPreference.menuAction())

        self.retranslateUi(2)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, lan):
        if lan == 1:
            self.MainWindow.setWindowTitle("Excel ??????????????????")
            self.MainWindow.setStatusTip("?????????")
            self.sourceLabel.setText("?????????")
            self.browseButton1.setText("??????")
            self.targetLabel.setText("????????????")
            self.browseButton2.setText("??????")
            self.headerNameLabel.setText("???????????????")
            self.startButton.setText("??????")
            self.logLabel.setText("??????")
            self.menuPreference.setTitle("??????")
            self.actionLanguage.setText("??????")
            self.actionClearLog.setText("????????????")
            self.actionAbout.setText("??????")
        else:
            self.MainWindow.setWindowTitle("Excel Auto Processor")
            self.MainWindow.setStatusTip("Welcome!")
            self.sourceLabel.setText("Source File")
            self.browseButton1.setText("Browse")
            self.targetLabel.setText("Target File")
            self.browseButton2.setText("Browse")
            self.headerNameLabel.setText("Key Column Header Name")
            self.startButton.setText("Start")
            self.logLabel.setText("Log")
            self.menuPreference.setTitle("Setting")
            self.actionLanguage.setText("Language")
            self.actionClearLog.setText("Clear Log")
            self.actionAbout.setText("About")

    def debugPrint( self, msgType, msg ):   #1:info  #2: warning  #3: error
        '''Print the message in the text edit at the bottom of the
        horizontal splitter.
        '''
        if msgType == 1:
            self.logBrowser.append( '<strong style="color:#43ba1e;">[INFO]</strong> ' + msg )    # Green
        elif msgType == 2:
            self.logBrowser.append( '<strong style="color:#aa7a19;">[WARNING]</strong> ' + msg )    # Yello
        else:
            self.logBrowser.append( '<strong style="color:#da3211;">[ERROR]</strong> ' + msg )      # Red
    
    def clearLog(self):
        # Clear log information in log browser
        self.logBrowser.clear()
    
    def openLanDialog(self):
        self.lanDialog.show()
    
    def openAboutDialog(self):
        self.aboutDialog.show()

    
    def browseButtonClicked(self, type):    # type = 1 or 2
        fileName = self.browseFile()
        if fileName:
            sheetNames = self.model.readSheetName(fileName, type)   # a sheet name set
            if type == 1:  
                self.debugPrint( 1, "Set source file name: " + fileName )
                self.lineEdit.setText(fileName)
                self.setComboBox(sheetNames, self.sheetSelect1)
            else:
                self.debugPrint( 1, "Set target file name: " + fileName )
                self.lineEdit_2.setText(fileName)
                self.setComboBox(sheetNames, self.sheetSelect2)

    
    def resetComboBox(self, comboBox):
        # delete all items in comboBox
        while comboBox.count() > 0:
            self.sheetSelect1.removeItem(0)


    def setComboBox(self, items, comboBox):    
        self.resetComboBox(comboBox)
        if len(items) > 1:
            comboBox.addItem('All sheets')
        for item in items:
            comboBox.addItem(str(item))

    
    def browseFile(self):
        # ???????????????????????????????????????
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, _ = QtWidgets.QFileDialog.getOpenFileName(
                        None,
                        "QFileDialog.getOpenFileName()",
                        "",
                        "WPS Excel?????? (*.xls);;Office Excel?????? (*.xlsx)",
                        options=options)
        return fileName
    
    def start(self):
        # Check if input is empty
        selectedSheet1 = self.sheetSelect1.currentText()
        if not selectedSheet1:
            self.debugPrint( 2, "Please select source file and its sheet" )
            return
        
        selectedSheet2 = self.sheetSelect2.currentText()
        if not selectedSheet2:
            self.debugPrint( 2, "Please select target file and its sheet" )
            return
        
        keyColumn = self.lineEdit_3.text()
        if not keyColumn:
            self.debugPrint( 2, "Please type in key column header name" )
            return

        self.model.process(keyColumn, selectedSheet1, selectedSheet2)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    '''
    Three style can choose: ['windowsvista', 'Windows', 'Fusion']
    '''

    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()

    sys.exit(app.exec_())
