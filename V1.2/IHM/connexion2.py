# -*- coding: utf-8 -*-

"""
Module implementing MainWindow.
"""

from PyQt4.QtCore import pyqtSlot
from PyQt4.QtGui import QMainWindow
from PyQt4.QtGui import QMessageBox


from GestionBdd import GestionBdd
from Menu import Menu

from Ui_connexion2 import Ui_MainWindow


class Connexion(QMainWindow, Ui_MainWindow):
    """
    Class documentation goes here.
    """
    def __init__(self, parent=None):
        """
        Constructor
        
        @param parent reference to the parent widget (QWidget)
        """
        super().__init__(parent)
        self.setupUi(self)
    
    @pyqtSlot()
    def on_buttonBox_2_accepted(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        # acces à la BDD
        db = GestionBdd('db')
        login = self.login.text()
        password = self.password.text()
        connexion = db.premiere_connexion(login, password)
        
        if connexion == True:            
            self.close()
            self.menu = Menu()
            self.menu.show()
        else:
#            self.emit(SIGNAL("Fermeture(PyQT_PyObject)"), connexion)
            QMessageBox.information(self, 
                    self.trUtf8("Erreur connexion "), 
                    self.trUtf8("Erreur sur le login et/ou mot de passe")) 

    
    @pyqtSlot()
    def on_buttonBox_2_rejected(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        self.close()
