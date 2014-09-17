# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'C:\Travail\EFS\Travail accreditation\SQ\Temperature\Logiciel Saisie des Etalonnages\Builds\V0.6\IHM\Menu.ui'
#
# Created: Sat May 17 00:47:42 2014
#      by: PyQt4 UI code generator 4.10
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

class Ui_Menu(object):
    def setupUi(self, Menu):
        Menu.setObjectName(_fromUtf8("Menu"))
        Menu.setEnabled(True)
        Menu.resize(770, 382)
        Menu.setMinimumSize(QtCore.QSize(770, 346))
        Menu.setMaximumSize(QtCore.QSize(770, 467))
        Menu.setAutoFillBackground(False)
        Menu.setStyleSheet(_fromUtf8("QGroupBox {\n"
"background-color: rgb(225, 225, 225); \n"
"border: 1px solid Black;\n"
"border-radius: 5px;\n"
" margin-top: 1ex; /* leave space at the top for the title */\n"
" }\n"
"\n"
"QRadioButton{\n"
"background-color: rgb(225, 225, 225);}\n"
"\n"
"\n"
"QComboBox {\n"
"background-color: rgb(225, 225, 225);\n"
"border: 1px solid Black;\n"
"border-radius: 5px;\n"
"padding: 1px 18px 1px 3px;}\n"
"\n"
"QLineEdit {\n"
"     border: 1px solid Black;\n"
"     border-radius: 5px;}\n"
"\n"
""))
        Menu.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly)
        Menu.setTabShape(QtGui.QTabWidget.Rounded)
        self.centralwidget = QtGui.QWidget(Menu)
        self.centralwidget.setObjectName(_fromUtf8("centralwidget"))
        self.groupBox = QtGui.QGroupBox(self.centralwidget)
        self.groupBox.setGeometry(QtCore.QRect(200, 30, 561, 251))
        font = QtGui.QFont()
        font.setFamily(_fromUtf8("Calibri"))
        font.setPointSize(12)
        self.groupBox.setFont(font)
        self.groupBox.setObjectName(_fromUtf8("groupBox"))
        self.gridLayoutWidget = QtGui.QWidget(self.groupBox)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(19, 30, 531, 191))
        self.gridLayoutWidget.setObjectName(_fromUtf8("gridLayoutWidget"))
        self.gridLayout = QtGui.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setMargin(0)
        self.gridLayout.setObjectName(_fromUtf8("gridLayout"))
        self.radioButton_saisie_etal = QtGui.QRadioButton(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.radioButton_saisie_etal.setFont(font)
        self.radioButton_saisie_etal.setObjectName(_fromUtf8("radioButton_saisie_etal"))
        self.gridLayout.addWidget(self.radioButton_saisie_etal, 0, 0, 1, 1)
        self.radioButton_validation = QtGui.QRadioButton(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.radioButton_validation.setFont(font)
        self.radioButton_validation.setObjectName(_fromUtf8("radioButton_validation"))
        self.gridLayout.addWidget(self.radioButton_validation, 1, 0, 1, 1)
        self.radioButton_consultation = QtGui.QRadioButton(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.radioButton_consultation.setFont(font)
        self.radioButton_consultation.setObjectName(_fromUtf8("radioButton_consultation"))
        self.gridLayout.addWidget(self.radioButton_consultation, 2, 0, 1, 1)
        self.comboBox_validation_n_doc = QtGui.QComboBox(self.gridLayoutWidget)
        self.comboBox_validation_n_doc.setEnabled(False)
        self.comboBox_validation_n_doc.setEditable(True)
        self.comboBox_validation_n_doc.setObjectName(_fromUtf8("comboBox_validation_n_doc"))
        self.gridLayout.addWidget(self.comboBox_validation_n_doc, 1, 1, 1, 1)
        self.comboBox_consultation = QtGui.QComboBox(self.gridLayoutWidget)
        self.comboBox_consultation.setEnabled(False)
        self.comboBox_consultation.setEditable(True)
        self.comboBox_consultation.setObjectName(_fromUtf8("comboBox_consultation"))
        self.gridLayout.addWidget(self.comboBox_consultation, 2, 1, 1, 1)
        self.radioButton_annule_remplace = QtGui.QRadioButton(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.radioButton_annule_remplace.setFont(font)
        self.radioButton_annule_remplace.setObjectName(_fromUtf8("radioButton_annule_remplace"))
        self.gridLayout.addWidget(self.radioButton_annule_remplace, 4, 0, 1, 1)
        self.radioButton_reeimpression = QtGui.QRadioButton(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.radioButton_reeimpression.setFont(font)
        self.radioButton_reeimpression.setObjectName(_fromUtf8("radioButton_reeimpression"))
        self.gridLayout.addWidget(self.radioButton_reeimpression, 3, 0, 1, 1)
        self.comboBox_annule_remplace_n_doc = QtGui.QComboBox(self.gridLayoutWidget)
        self.comboBox_annule_remplace_n_doc.setEnabled(False)
        self.comboBox_annule_remplace_n_doc.setEditable(True)
        self.comboBox_annule_remplace_n_doc.setObjectName(_fromUtf8("comboBox_annule_remplace_n_doc"))
        self.gridLayout.addWidget(self.comboBox_annule_remplace_n_doc, 4, 1, 1, 1)
        self.comboBox_reeimp_n_document = QtGui.QComboBox(self.gridLayoutWidget)
        self.comboBox_reeimp_n_document.setEnabled(False)
        self.comboBox_reeimp_n_document.setEditable(True)
        self.comboBox_reeimp_n_document.setObjectName(_fromUtf8("comboBox_reeimp_n_document"))
        self.gridLayout.addWidget(self.comboBox_reeimp_n_document, 3, 1, 1, 1)
        self.label = QtGui.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(20, 10, 171, 201))
        self.label.setText(_fromUtf8(""))
        self.label.setPixmap(QtGui.QPixmap(_fromUtf8("LOGO_REACTUALISE_petit_format.png")))
        self.label.setObjectName(_fromUtf8("label"))
        Menu.setCentralWidget(self.centralwidget)
        self.menubar = QtGui.QMenuBar(Menu)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 770, 21))
        self.menubar.setObjectName(_fromUtf8("menubar"))
        Menu.setMenuBar(self.menubar)
        self.statusbar = QtGui.QStatusBar(Menu)
        self.statusbar.setObjectName(_fromUtf8("statusbar"))
        Menu.setStatusBar(self.statusbar)

        self.retranslateUi(Menu)
        QtCore.QObject.connect(self.radioButton_saisie_etal, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_validation_n_doc.setDisabled)
        QtCore.QObject.connect(self.radioButton_saisie_etal, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_validation_n_doc.clear)
        QtCore.QObject.connect(self.radioButton_saisie_etal, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_consultation.setDisabled)
        QtCore.QObject.connect(self.radioButton_saisie_etal, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_consultation.clear)
        QtCore.QObject.connect(self.radioButton_saisie_etal, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_reeimp_n_document.setDisabled)
        QtCore.QObject.connect(self.radioButton_saisie_etal, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_reeimp_n_document.clear)
        QtCore.QObject.connect(self.radioButton_saisie_etal, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_annule_remplace_n_doc.setDisabled)
        QtCore.QObject.connect(self.radioButton_saisie_etal, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_annule_remplace_n_doc.clear)
        QtCore.QObject.connect(self.radioButton_validation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_validation_n_doc.setEnabled)
        QtCore.QObject.connect(self.radioButton_validation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_validation_n_doc.clear)
        QtCore.QObject.connect(self.radioButton_validation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_consultation.setDisabled)
        QtCore.QObject.connect(self.radioButton_validation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_consultation.clear)
        QtCore.QObject.connect(self.radioButton_validation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_reeimp_n_document.setDisabled)
        QtCore.QObject.connect(self.radioButton_validation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_reeimp_n_document.clear)
        QtCore.QObject.connect(self.radioButton_validation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_annule_remplace_n_doc.setDisabled)
        QtCore.QObject.connect(self.radioButton_validation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_annule_remplace_n_doc.clear)
        QtCore.QObject.connect(self.radioButton_consultation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_consultation.setEnabled)
        QtCore.QObject.connect(self.radioButton_consultation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_consultation.clear)
        QtCore.QObject.connect(self.radioButton_consultation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_validation_n_doc.setDisabled)
        QtCore.QObject.connect(self.radioButton_consultation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_validation_n_doc.clear)
        QtCore.QObject.connect(self.radioButton_consultation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_reeimp_n_document.setDisabled)
        QtCore.QObject.connect(self.radioButton_consultation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_reeimp_n_document.clear)
        QtCore.QObject.connect(self.radioButton_consultation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_annule_remplace_n_doc.setDisabled)
        QtCore.QObject.connect(self.radioButton_consultation, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_annule_remplace_n_doc.clear)
        QtCore.QObject.connect(self.radioButton_reeimpression, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_reeimp_n_document.setEnabled)
        QtCore.QObject.connect(self.radioButton_reeimpression, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_reeimp_n_document.clear)
        QtCore.QObject.connect(self.radioButton_reeimpression, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_validation_n_doc.setDisabled)
        QtCore.QObject.connect(self.radioButton_reeimpression, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_validation_n_doc.clear)
        QtCore.QObject.connect(self.radioButton_reeimpression, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_consultation.setDisabled)
        QtCore.QObject.connect(self.radioButton_reeimpression, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_consultation.clear)
        QtCore.QObject.connect(self.radioButton_reeimpression, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_annule_remplace_n_doc.setDisabled)
        QtCore.QObject.connect(self.radioButton_reeimpression, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_annule_remplace_n_doc.clear)
        QtCore.QObject.connect(self.radioButton_annule_remplace, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_annule_remplace_n_doc.setEnabled)
        QtCore.QObject.connect(self.radioButton_annule_remplace, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_annule_remplace_n_doc.clear)
        QtCore.QObject.connect(self.radioButton_annule_remplace, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_validation_n_doc.setDisabled)
        QtCore.QObject.connect(self.radioButton_annule_remplace, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_validation_n_doc.clear)
        QtCore.QObject.connect(self.radioButton_annule_remplace, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_consultation.setDisabled)
        QtCore.QObject.connect(self.radioButton_annule_remplace, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_consultation.clear)
        QtCore.QObject.connect(self.radioButton_annule_remplace, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_reeimp_n_document.setDisabled)
        QtCore.QObject.connect(self.radioButton_annule_remplace, QtCore.SIGNAL(_fromUtf8("clicked(bool)")), self.comboBox_reeimp_n_document.clear)
        QtCore.QMetaObject.connectSlotsByName(Menu)

    def retranslateUi(self, Menu):
        Menu.setWindowTitle(_translate("Menu", "Menu Etalonnage en Temperature", None))
        Menu.setWhatsThis(_translate("Menu", "<html><head/><body><p><img src=\":/newPrefix/LOGO REACTUALISE.bmp\"/></p></body></html>", None))
        self.groupBox.setTitle(_translate("Menu", "Menu", None))
        self.radioButton_saisie_etal.setText(_translate("Menu", "Saisie Etalonnage ", None))
        self.radioButton_validation.setText(_translate("Menu", "Validation / Modification", None))
        self.radioButton_consultation.setText(_translate("Menu", "Consultation saisies", None))
        self.radioButton_annule_remplace.setText(_translate("Menu", "Annule et Remplace", None))
        self.radioButton_reeimpression.setText(_translate("Menu", "Reimpression :  n° du document", None))


if __name__ == "__main__":
    import sys
    app = QtGui.QApplication(sys.argv)
    Menu = QtGui.QMainWindow()
    ui = Ui_Menu()
    ui.setupUi(Menu)
    Menu.show()
    sys.exit(app.exec_())

