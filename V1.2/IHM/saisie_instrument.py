# -*- coding: utf-8 -*-

"""
Module implementing SaisieInstrument.
"""
import sys
from PyQt4 import QtGui
from PyQt4.QtCore import pyqtSlot
from PyQt4.QtGui import QMainWindow

from Ui_saisie_instrument import Ui_SaisieInstrument


class SaisieInstrument(QMainWindow, Ui_SaisieInstrument):
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
    def on_textEdit_n_serie_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_textEdit_sap_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_textEdit_identification_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_textEdit_commentaire_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_textEdit_domaine_mesure_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        print("dans ton ul")
    
    @pyqtSlot()
    def on_textEdit_famille_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_textEdit_particularite_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_textEdit_resolution_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    

    
    @pyqtSlot()
    def on_textEdit_duree_periodicite_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_textEdit_gestionnaire_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_textEdit_localisation_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_textEdit_site_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_textEdit_sous_affectation_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_textEdit_etat_utilisation_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_textEdit_prestataire_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_textEdit_periodicite_cursorPositionChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot(str)
    def on_comboBox_ref_constructeur_currentIndexChanged(self, p0):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        print("hello")
    
    @pyqtSlot(str)
    def on_comboBox_constructeur_currentIndexChanged(self, p0):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot(str)
    def on_comboBox_designation_currentIndexChanged(self, p0):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot(str)
    def on_comboBox_statut_currentIndexChanged(self, p0):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_buttonBox_accepted(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_buttonBox_rejected(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        
    
    @pyqtSlot()
    def on_textEdit_code_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        print("coucuoc")

    

        
    @pyqtSlot(int)
    def on_comboBox_constructeur_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        print("je mactive")
    
    @pyqtSlot(int)
    def on_comboBox_designation_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot(int)
    def on_comboBox_statut_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot(str)
    def on_comboBox_ref_constructeur_activated(self, p0):
        """
        Slot documentation goes here.
        """
        print("je mactive")    
        
if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)

    main = SaisieInstrument()
    main.show()

    sys.exit(app.exec_())
