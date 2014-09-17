# -*- coding: utf-8 -*-

"""
Module implementing Form.
"""

from PyQt4.QtCore import pyqtSlot
from PyQt4.QtGui import QWidget
from PyQt4 import QtGui, QtCore

from Ui_resume_onglet_resultat import Ui_Form


class Form(QWidget, Ui_Form):
    """
    Class documentation goes here.
    """
    def __init__(self, nom_instrum, nbr_ligne, donnees, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget (QWidget)
        """
        super().__init__(parent)
        self.setupUi(self)       
        self.label.setText(nom_instrum)

        i=0
        while i<=4:            
            self.tableWidget.setColumnWidth(i,150)
            
            i+=1
            
        for j in range(nbr_ligne):            
            self.tableWidget.insertRow(0)
            j+=1
            
        i=0
        while i < nbr_ligne:
            
            self.tableWidget.setItem(i, 0, QtGui.QTableWidgetItem(str(donnees["moyennes_etalon_non_corrigees"][i])))
            self.tableWidget.setItem(i, 1, QtGui.QTableWidgetItem(str(donnees["moyennes_etalon_corrigees"][i])))
            self.tableWidget.setItem(i, 2, QtGui.QTableWidgetItem(str(donnees["moyennes_instrument"][i])))   
            self.tableWidget.setItem(i, 3, QtGui.QTableWidgetItem(str(donnees["corrections"][i])))
            self.tableWidget.setItem(i, 4, QtGui.QTableWidgetItem(str(donnees["incertitudes"][i])))  
                
            i+=1
        

            
    @pyqtSlot()
    def on_pushButton_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        self.close()
