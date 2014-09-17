# -*- coding: utf-8 -*-

"""
Module implementing MainWindow.
"""
import os
import numpy
from PyQt4 import QtGui, QtCore
from PyQt4.QtCore import pyqtSlot
from PyQt4.QtGui import QMainWindow
from PyQt4.QtGui import QFileDialog
from PyQt4.QtGui import QMessageBox
from PyQt4.QtGui import QInputDialog
#from PyQt4 import  QtSql
from Ui_Saisie_etalonnage import Ui_Saisie
from Package.fonctions import traitement_fichier_etalon
from Package.fonctions import  mise_enforme_saisies
from Package.fonctions import sauvegarde_onglet_saisie
from Package.fonctions import lecture_onglet_saisie
from Package.fonctions import calcul_polynome
import decimal
from RapportEtalonnage import RapportEtalonnage
from OngletConformite import OngletConformite
from GestionBdd import GestionBdd
import datetime
from resume_onglet_resultat import Form


class Saisie(QMainWindow, Ui_Saisie):
    """
    Class documentation goes here.
    """
    def __init__(self, type_de_saisie, parent=None):
        """Constructor        
        @param parent reference to the parent widget (QWidget)
        type_saisie est une liste :
        type_saisie = [1] : saisie d'un etalonnage normale
        type_saisie = [2,n_campagne] : modification d'un etalonnage (mise à jour)
        type_saisie = [3,n_certificat] : Annule et remplace n_certificat
        """
        super().__init__(parent)
        self.setupUi(self)
        self.type_de_saisie = type_de_saisie # si modif = 1 ca veut dire qu'on est en saisie si 2 on est en modif /annule remplace
       
        # acces à la BDD
        db = GestionBdd('db')
        db.reconnexion()
          
        # Gestion des combobox:        
#        list_identification_instrum = db.gestion_combobox_onglet_config()
        list_identification_instrum = db.gestion_combobox_onglet_config_2()
               
        self.comboBox_ident_instrum_1.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_2.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_3.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_4.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_5.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_6.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_7.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_8.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_9.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_10.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_11.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_12.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_13.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_14.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_15.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_16.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_17.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_18.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_19.addItems(list_identification_instrum)
        self.comboBox_ident_instrum_20.addItems(list_identification_instrum)
        
        

       
       #gestion combobox operateur

        list_technicien = db.gestion_combobox_onglet_operateur_2()
        self.Combobox_operateur_select_2.addItems(list_technicien)
      
        #Gestion des spinboxs : nbr_temperature et nbr d'instruments
        self.SpinBox_nbr_instruments.setMinimum(1)
        self.SpinBox_nbr_instruments.setMaximum(20)
        
        self. SpinBox_nbr_pt_etal.setMinimum(1)
        self.SpinBox_nbr_releves_2.setValue(10)
        
        #interdit l'ecriture sur les onglet saisie resultats 
        self.tab_2.setEnabled(False)
        self.tab_3.setEnabled(False)
        self.actionEnregistrer.setEnabled(False)
        
        #Gestion combobox onglet conformité        
        #Acces bdd
#        db = GestionBdd('db')
#        db.reconnexion()
        list_emt = db.recuperation_ref_emt()
        
        self.comboBox_ref_emt_1.addItems(list_emt)
        self.comboBox_ref_emt_2.addItems(list_emt)
        self.comboBox_ref_emt_3.addItems(list_emt)
        self.comboBox_ref_emt_4.addItems(list_emt)
        self.comboBox_ref_emt_5.addItems(list_emt)
        self.comboBox_ref_emt_6.addItems(list_emt)
        self.comboBox_ref_emt_7.addItems(list_emt)
        self.comboBox_ref_emt_8.addItems(list_emt)
        self.comboBox_ref_emt_9.addItems(list_emt)
        self.comboBox_ref_emt_10.addItems(list_emt)
        self.comboBox_ref_emt_11.addItems(list_emt)
        self.comboBox_ref_emt_12.addItems(list_emt)
        self.comboBox_ref_emt_13.addItems(list_emt)
        self.comboBox_ref_emt_14.addItems(list_emt)
        self.comboBox_ref_emt_15.addItems(list_emt)
        self.comboBox_ref_emt_16.addItems(list_emt)
        self.comboBox_ref_emt_17.addItems(list_emt)
        self.comboBox_ref_emt_18.addItems(list_emt)
        self.comboBox_ref_emt_19.addItems(list_emt)
        self.comboBox_ref_emt_20.addItems(list_emt)
        
        list_text_edit_etat_reception = [self.textEdit_etat_reception_inst_1, self.textEdit_etat_reception_inst_2, self.textEdit_etat_reception_inst_3,
                                                                self.textEdit_etat_reception_inst_4, self.textEdit_etat_reception_inst_5, self.textEdit_etat_reception_inst_6, 
                                                                self.textEdit_etat_reception_inst_7, self.textEdit_etat_reception_inst_8, self.textEdit_etat_reception_inst_9,
                                                                self.textEdit_etat_reception_inst_10, self.textEdit_etat_reception_inst_11, self.textEdit_etat_reception_inst_12,
                                                                self.textEdit_etat_reception_inst_13, self.textEdit_etat_reception_inst_14, self.textEdit_etat_reception_inst_15,
                                                                self.textEdit_etat_reception_inst_16, self.textEdit_etat_reception_inst_17, self.textEdit_etat_reception_inst_18,
                                                                self.textEdit_etat_reception_inst_19, self.textEdit_etat_reception_inst_20]
        
        #Verification qu'un fichier de configuration n'existe pas s'il existe on affiche
        path =os.path.abspath("AppData/configuration") #permet de recuperer le chemin absolu du fichier le test ne marche pas sinon
        
        
        if self.type_de_saisie[0] == 1: #si c'est une saisie normale on demande à l'utilisateur sil souhaite importer la campagne
            
            if os.path.isfile(path) == True :
                reponse = QMessageBox.question(self, 
                    self.trUtf8("Information"), 
                    self.trUtf8("Une configuration d'etalonnage est archivée voulez vous l'importer"), 
                    QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
        
            
                if reponse == QtGui.QMessageBox.Yes :
                                       
                    onglet = lecture_onglet_saisie("configuration")
                    
                    self.SpinBox_nbr_instruments.setValue(onglet[0]["nbr_instrum"])
                    self.SpinBox_nbr_pt_etal.setValue(onglet[0]["nbr_pts_temp"])
                    self.calendarWidget.setSelectedDate(onglet[0]["Date"])
                    
                    list_combobox_configuration = [self.comboBox_ident_instrum_1, self.comboBox_ident_instrum_2, self.comboBox_ident_instrum_3, 
                                                                self.comboBox_ident_instrum_4, self.comboBox_ident_instrum_5, self.comboBox_ident_instrum_6, 
                                                                self.comboBox_ident_instrum_7, self.comboBox_ident_instrum_8, self.comboBox_ident_instrum_9, 
                                                                self.comboBox_ident_instrum_10, self.comboBox_ident_instrum_11, self.comboBox_ident_instrum_12, 
                                                                self.comboBox_ident_instrum_13, self.comboBox_ident_instrum_14, self.comboBox_ident_instrum_15, 
                                                                self.comboBox_ident_instrum_16, self.comboBox_ident_instrum_17, self.comboBox_ident_instrum_18, 
                                                                self.comboBox_ident_instrum_19, self.comboBox_ident_instrum_20]
                    
                    
                    list_textedit_configuration = [self.textEdit_n_serie_instrum_1, self.textEdit_construct_instrum_1, self.textEdit_type_instrum_1, self.textEdit_resolution_instrum_1, self.textEdit_prof_imm_instrum_1, self.radioButton_cofrac_1, self.radioButton_non_cofrac_1, self.textEdit_com_inst_1, 
                                                           self.textEdit_n_serie_instrum_2, self.textEdit_construct_instrum_2, self.textEdit_type_instrum_2, self.textEdit_resolution_instrum_2, self.textEdit_prof_imm_instrum_2, self.radioButton_cofrac_2, self.radioButton_non_cofrac_2, self.textEdit_com_inst_2,
                                                           self.textEdit_n_serie_instrum_3, self.textEdit_construct_instrum_3, self.textEdit_type_instrum_3, self.textEdit_resolution_instrum_3, self.textEdit_prof_imm_instrum_3, self.radioButton_cofrac_3, self.radioButton_non_cofrac_3, self.textEdit_com_inst_3,
                                                           self.textEdit_n_serie_instrum_4, self.textEdit_construct_instrum_4, self.textEdit_type_instrum_4, self.textEdit_resolution_instrum_4, self.textEdit_prof_imm_instrum_4, self.radioButton_cofrac_4, self.radioButton_non_cofrac_4, self.textEdit_com_inst_4,
                                                           self.textEdit_n_serie_instrum_5, self.textEdit_construct_instrum_5, self.textEdit_type_instrum_5, self.textEdit_resolution_instrum_5, self.textEdit_prof_imm_instrum_5, self.radioButton_cofrac_5, self.radioButton_non_cofrac_5, self.textEdit_com_inst_5,
                                                           self.textEdit_n_serie_instrum_6, self.textEdit_construct_instrum_6, self.textEdit_type_instrum_6, self.textEdit_resolution_instrum_6, self.textEdit_prof_imm_instrum_6, self.radioButton_cofrac_6, self.radioButton_non_cofrac_6, self.textEdit_com_inst_6,
                                                           self.textEdit_n_serie_instrum_7, self.textEdit_construct_instrum_7, self.textEdit_type_instrum_7, self.textEdit_resolution_instrum_7, self.textEdit_prof_imm_instrum_7, self.radioButton_cofrac_7, self.radioButton_non_cofrac_7, self.textEdit_com_inst_7,
                                                           self.textEdit_n_serie_instrum_8, self.textEdit_construct_instrum_8, self.textEdit_type_instrum_8, self.textEdit_resolution_instrum_8, self.textEdit_prof_imm_instrum_8, self.radioButton_cofrac_8, self.radioButton_non_cofrac_8, self.textEdit_com_inst_8,
                                                           self.textEdit_n_serie_instrum_9, self.textEdit_construct_instrum_9, self.textEdit_type_instrum_9, self.textEdit_resolution_instrum_9, self.textEdit_prof_imm_instrum_9, self.radioButton_cofrac_9, self.radioButton_non_cofrac_9, self.textEdit_com_inst_9,
                                                          self.textEdit_n_serie_instrum_10, self.textEdit_construct_instrum_10, self.textEdit_type_instrum_10, self.textEdit_resolution_instrum_10, self.textEdit_prof_imm_instrum_10, self.radioButton_cofrac_10, self.radioButton_non_cofrac_10, self.textEdit_com_inst_10,
                                                          self.textEdit_n_serie_instrum_11, self.textEdit_construct_instrum_11, self.textEdit_type_instrum_11, self.textEdit_resolution_instrum_11, self.textEdit_prof_imm_instrum_11, self.radioButton_cofrac_11, self.radioButton_non_cofrac_11, self.textEdit_com_inst_11,
                                                          self.textEdit_n_serie_instrum_12, self.textEdit_construct_instrum_12, self.textEdit_type_instrum_12, self.textEdit_resolution_instrum_12, self.textEdit_prof_imm_instrum_12, self.radioButton_cofrac_12, self.radioButton_non_cofrac_12, self.textEdit_com_inst_12,
                                                          self.textEdit_n_serie_instrum_13, self.textEdit_construct_instrum_13, self.textEdit_type_instrum_13, self.textEdit_resolution_instrum_13, self.textEdit_prof_imm_instrum_13, self.radioButton_cofrac_13, self.radioButton_non_cofrac_13, self.textEdit_com_inst_13,
                                                          self.textEdit_n_serie_instrum_14, self.textEdit_construct_instrum_14, self.textEdit_type_instrum_14, self.textEdit_resolution_instrum_14, self.textEdit_prof_imm_instrum_14, self.radioButton_cofrac_14, self.radioButton_non_cofrac_14, self.textEdit_com_inst_14,
                                                          self.textEdit_n_serie_instrum_15, self.textEdit_construct_instrum_15, self.textEdit_type_instrum_15, self.textEdit_resolution_instrum_15, self.textEdit_prof_imm_instrum_15, self.radioButton_cofrac_15, self.radioButton_non_cofrac_15, self.textEdit_com_inst_15,
                                                          self.textEdit_n_serie_instrum_16, self.textEdit_construct_instrum_16, self.textEdit_type_instrum_16, self.textEdit_resolution_instrum_16, self.textEdit_prof_imm_instrum_16, self.radioButton_cofrac_16, self.radioButton_non_cofrac_16, self.textEdit_com_inst_16,
                                                          self.textEdit_n_serie_instrum_17, self.textEdit_construct_instrum_17, self.textEdit_type_instrum_17, self.textEdit_resolution_instrum_17, self.textEdit_prof_imm_instrum_17, self.radioButton_cofrac_17, self.radioButton_non_cofrac_17, self.textEdit_com_inst_17,
                                                          self.textEdit_n_serie_instrum_18, self.textEdit_construct_instrum_18, self.textEdit_type_instrum_18, self.textEdit_resolution_instrum_18, self.textEdit_prof_imm_instrum_18, self.radioButton_cofrac_18, self.radioButton_non_cofrac_18, self.textEdit_com_inst_18,
                                                          self.textEdit_n_serie_instrum_19, self.textEdit_construct_instrum_19, self.textEdit_type_instrum_19, self.textEdit_resolution_instrum_19, self.textEdit_prof_imm_instrum_19, self.radioButton_cofrac_19, self.radioButton_non_cofrac_19, self.textEdit_com_inst_19,
                                                          self.textEdit_n_serie_instrum_20, self.textEdit_construct_instrum_20, self.textEdit_type_instrum_20, self.textEdit_resolution_instrum_20, self.textEdit_prof_imm_instrum_20, self.radioButton_cofrac_20, self.radioButton_non_cofrac_20, self.textEdit_com_inst_20
                                                          ]
                    
                   
                    i = 0
                    while i < onglet[0]["nbr_instrum"] :
                       
                       #on recupere l'index correspondant au nom de l'instrument 
                      
                        index = list_combobox_configuration[i].findText(onglet[i+1]["nom_instrum"])
                        list_combobox_configuration[i].setCurrentIndex(int(index))
                        list_textedit_configuration[(8*i )].append(str(onglet[i+1]["n_serie"]))
                        list_textedit_configuration[(8*i + 1)].append(onglet[i+1]["constructeur"])
                        list_textedit_configuration[(8*i + 2)].append(onglet[i+1]["Type"])
                        
                        for ele in onglet[i+1]["resolution"] :
                            list_textedit_configuration[(8*i + 3)].append(str(ele))
                        for ele in onglet[i+1]["immersion"] :
                            list_textedit_configuration[(8*i + 4)].append(str(ele))                   
                        
                        if onglet[i+1]["Type_etalonnage"] == "Cofrac" :
                            list_textedit_configuration[(8*i + 5)].setChecked(True)
                        else :
                            list_textedit_configuration[(8*i + 6)].setChecked(True)  
                        
                        list_textedit_configuration[(8*i + 7)].append(onglet[i+1]["renseignement_complementaire"])
                        
                        list_text_edit_etat_reception[i].clear()
                        list_text_edit_etat_reception[i].append(onglet[i+1]["Etat_reception"])
                        
                        i+=1
    
                else :
                    #effacement des fichiers temporaire (present dans le dossier AppData)
                    path =os.path.abspath("AppData/")
                    
                    for ele in os.listdir(path):                    
                        path_total = str(path + "/"+str(ele))            
                        if os.path.isfile(path_total): # verification qu'il s'agit bien de fichier
                            os.remove(path_total)

        else:
            onglet = lecture_onglet_saisie("configuration")
          
                    
            self.SpinBox_nbr_instruments.setValue(onglet[0]["nbr_instrum"])
            self.SpinBox_nbr_pt_etal.setValue(onglet[0]["nbr_pts_temp"])
            self.calendarWidget.setSelectedDate(onglet[0]["Date"])
                    
            list_combobox_configuration = [self.comboBox_ident_instrum_1, self.comboBox_ident_instrum_2, self.comboBox_ident_instrum_3, 
                                                        self.comboBox_ident_instrum_4, self.comboBox_ident_instrum_5, self.comboBox_ident_instrum_6, 
                                                        self.comboBox_ident_instrum_7, self.comboBox_ident_instrum_8, self.comboBox_ident_instrum_9, 
                                                        self.comboBox_ident_instrum_10, self.comboBox_ident_instrum_11, self.comboBox_ident_instrum_12, 
                                                        self.comboBox_ident_instrum_13, self.comboBox_ident_instrum_14, self.comboBox_ident_instrum_15, 
                                                        self.comboBox_ident_instrum_16, self.comboBox_ident_instrum_17, self.comboBox_ident_instrum_18, 
                                                        self.comboBox_ident_instrum_19, self.comboBox_ident_instrum_20]
            
            
            list_textedit_configuration = [self.textEdit_n_serie_instrum_1, self.textEdit_construct_instrum_1, self.textEdit_type_instrum_1, self.textEdit_resolution_instrum_1, self.textEdit_prof_imm_instrum_1, self.radioButton_cofrac_1, self.radioButton_non_cofrac_1, self.textEdit_com_inst_1, 
                                                   self.textEdit_n_serie_instrum_2, self.textEdit_construct_instrum_2, self.textEdit_type_instrum_2, self.textEdit_resolution_instrum_2, self.textEdit_prof_imm_instrum_2, self.radioButton_cofrac_2, self.radioButton_non_cofrac_2, self.textEdit_com_inst_2,
                                                   self.textEdit_n_serie_instrum_3, self.textEdit_construct_instrum_3, self.textEdit_type_instrum_3, self.textEdit_resolution_instrum_3, self.textEdit_prof_imm_instrum_3, self.radioButton_cofrac_3, self.radioButton_non_cofrac_3, self.textEdit_com_inst_3,
                                                   self.textEdit_n_serie_instrum_4, self.textEdit_construct_instrum_4, self.textEdit_type_instrum_4, self.textEdit_resolution_instrum_4, self.textEdit_prof_imm_instrum_4, self.radioButton_cofrac_4, self.radioButton_non_cofrac_4, self.textEdit_com_inst_4,
                                                   self.textEdit_n_serie_instrum_5, self.textEdit_construct_instrum_5, self.textEdit_type_instrum_5, self.textEdit_resolution_instrum_5, self.textEdit_prof_imm_instrum_5, self.radioButton_cofrac_5, self.radioButton_non_cofrac_5, self.textEdit_com_inst_5,
                                                   self.textEdit_n_serie_instrum_6, self.textEdit_construct_instrum_6, self.textEdit_type_instrum_6, self.textEdit_resolution_instrum_6, self.textEdit_prof_imm_instrum_6, self.radioButton_cofrac_6, self.radioButton_non_cofrac_6, self.textEdit_com_inst_6,
                                                   self.textEdit_n_serie_instrum_7, self.textEdit_construct_instrum_7, self.textEdit_type_instrum_7, self.textEdit_resolution_instrum_7, self.textEdit_prof_imm_instrum_7, self.radioButton_cofrac_7, self.radioButton_non_cofrac_7, self.textEdit_com_inst_7,
                                                   self.textEdit_n_serie_instrum_8, self.textEdit_construct_instrum_8, self.textEdit_type_instrum_8, self.textEdit_resolution_instrum_8, self.textEdit_prof_imm_instrum_8, self.radioButton_cofrac_8, self.radioButton_non_cofrac_8, self.textEdit_com_inst_8,
                                                   self.textEdit_n_serie_instrum_9, self.textEdit_construct_instrum_9, self.textEdit_type_instrum_9, self.textEdit_resolution_instrum_9, self.textEdit_prof_imm_instrum_9, self.radioButton_cofrac_9, self.radioButton_non_cofrac_9, self.textEdit_com_inst_9,
                                                  self.textEdit_n_serie_instrum_10, self.textEdit_construct_instrum_10, self.textEdit_type_instrum_10, self.textEdit_resolution_instrum_10, self.textEdit_prof_imm_instrum_10, self.radioButton_cofrac_10, self.radioButton_non_cofrac_10, self.textEdit_com_inst_10,
                                                  self.textEdit_n_serie_instrum_11, self.textEdit_construct_instrum_11, self.textEdit_type_instrum_11, self.textEdit_resolution_instrum_11, self.textEdit_prof_imm_instrum_11, self.radioButton_cofrac_11, self.radioButton_non_cofrac_11, self.textEdit_com_inst_11,
                                                  self.textEdit_n_serie_instrum_12, self.textEdit_construct_instrum_12, self.textEdit_type_instrum_12, self.textEdit_resolution_instrum_12, self.textEdit_prof_imm_instrum_12, self.radioButton_cofrac_12, self.radioButton_non_cofrac_12, self.textEdit_com_inst_12,
                                                  self.textEdit_n_serie_instrum_13, self.textEdit_construct_instrum_13, self.textEdit_type_instrum_13, self.textEdit_resolution_instrum_13, self.textEdit_prof_imm_instrum_13, self.radioButton_cofrac_13, self.radioButton_non_cofrac_13, self.textEdit_com_inst_13,
                                                  self.textEdit_n_serie_instrum_14, self.textEdit_construct_instrum_14, self.textEdit_type_instrum_14, self.textEdit_resolution_instrum_14, self.textEdit_prof_imm_instrum_14, self.radioButton_cofrac_14, self.radioButton_non_cofrac_14, self.textEdit_com_inst_14,
                                                  self.textEdit_n_serie_instrum_15, self.textEdit_construct_instrum_15, self.textEdit_type_instrum_15, self.textEdit_resolution_instrum_15, self.textEdit_prof_imm_instrum_15, self.radioButton_cofrac_15, self.radioButton_non_cofrac_15, self.textEdit_com_inst_15,
                                                  self.textEdit_n_serie_instrum_16, self.textEdit_construct_instrum_16, self.textEdit_type_instrum_16, self.textEdit_resolution_instrum_16, self.textEdit_prof_imm_instrum_16, self.radioButton_cofrac_16, self.radioButton_non_cofrac_16, self.textEdit_com_inst_16,
                                                  self.textEdit_n_serie_instrum_17, self.textEdit_construct_instrum_17, self.textEdit_type_instrum_17, self.textEdit_resolution_instrum_17, self.textEdit_prof_imm_instrum_17, self.radioButton_cofrac_17, self.radioButton_non_cofrac_17, self.textEdit_com_inst_17,
                                                  self.textEdit_n_serie_instrum_18, self.textEdit_construct_instrum_18, self.textEdit_type_instrum_18, self.textEdit_resolution_instrum_18, self.textEdit_prof_imm_instrum_18, self.radioButton_cofrac_18, self.radioButton_non_cofrac_18, self.textEdit_com_inst_18,
                                                  self.textEdit_n_serie_instrum_19, self.textEdit_construct_instrum_19, self.textEdit_type_instrum_19, self.textEdit_resolution_instrum_19, self.textEdit_prof_imm_instrum_19, self.radioButton_cofrac_19, self.radioButton_non_cofrac_19, self.textEdit_com_inst_19,
                                                  self.textEdit_n_serie_instrum_20, self.textEdit_construct_instrum_20, self.textEdit_type_instrum_20, self.textEdit_resolution_instrum_20, self.textEdit_prof_imm_instrum_20, self.radioButton_cofrac_20, self.radioButton_non_cofrac_20, self.textEdit_com_inst_20
                                                  ]
       
           
            i = 0
            while i < onglet[0]["nbr_instrum"] :
               
               #on recupere l'index correspondant au nom de l'instrument 
              
                index = list_combobox_configuration[i].findText(onglet[i+1]["nom_instrum"])
                list_combobox_configuration[i].setCurrentIndex(int(index))
                list_textedit_configuration[(8*i )].append(str(onglet[i+1]["n_serie"]))
                list_textedit_configuration[(8*i + 1)].append(onglet[i+1]["constructeur"])
                list_textedit_configuration[(8*i + 2)].append(onglet[i+1]["Type"])
                
                for ele in onglet[i+1]["resolution"] :
                    list_textedit_configuration[(8*i + 3)].append(str(ele))
                for ele in onglet[i+1]["immersion"] :
                    list_textedit_configuration[(8*i + 4)].append(str(ele))                   
                
                if onglet[i+1]["Type_etalonnage"] == "Cofrac" :
                    list_textedit_configuration[(8*i + 5)].setChecked(True)
                else :
                    list_textedit_configuration[(8*i + 6)].setChecked(True)  
                
                list_textedit_configuration[(8*i + 7)].append(onglet[i+1]["renseignement_complementaire"])
                
                list_text_edit_etat_reception[i].clear()
                list_text_edit_etat_reception[i].append(onglet[i+1]["Etat_reception"])
                
                i+=1
                

    
    @pyqtSlot()
    def on_Bouton_fichier_etalon_2_clicked(self):
        '''gestion de l'appui sur le bouton selection donneees etalon .Importation de ces donnees'''
        '''Permet d'utiliser l'explorateur systeme  avec qt et selectionner un fichier'''
        
        try :
            self.Zone_texte_chemin_fichier_etalon_2.clear()
            self.textEdit_mesures_etalon_brutes.clear()
            path = 'y:/1.METROLOGIE/0.ARCHIVES ETALONNAGE-VERIFICATIONS/1-TEMPERATURE/'#os.getcwd()
            fichier = QFileDialog.getOpenFileName(self, "Ouvrir un fichier", path, "Texte(*.txt)")
            self.Zone_texte_chemin_fichier_etalon_2.append(fichier)
            
            #appel de la fct traitement_fichier_etalon qui renvoie un dict avec en entree l'indice correspond a la colonne d'une donnée et en valeur
            #une liste avec les donnees correspondantes 
            
            
            donnees = traitement_fichier_etalon(fichier)
    #        
            #on recherche le nbr d'indice present dans le dictionnaire
            
            nbr_indice_donnees  = -1 # -1 car le dictionnaire commence à 0
           
            for clef in donnees.keys():
                nbr_indice_donnees += 1
           
           
            for elements in donnees[2] :
                self.textEdit_mesures_etalon_brutes.append(elements) 
            
            #on test voir si ces indices sont dans le dictionnaire (attention max 6 etalons)
            # rq à ameliorer 
            if nbr_indice_donnees >= 4:
                for elements in donnees[4] :
                    self.textEdit_mesures_inst_1.append(elements)
            if nbr_indice_donnees >= 6:
                for elements in donnees[6] :
                    self.textEdit_mesures_inst_2.append(elements)
            if nbr_indice_donnees >= 8:
                for elements in donnees[8] :
                    self.textEdit_mesures_inst_3.append(elements)       
            
            if nbr_indice_donnees >= 10 :
                for elements in donnees[10] :
                    self.textEdit_mesures_inst_4.append(elements)
            
            if nbr_indice_donnees >= 12 :
                for elements in donnees[12] :
                    self.textEdit_mesures_inst_5.append(elements)       
           
            #return filename
        except AttributeError :
            pass
            
            
    @pyqtSlot()
    def on_pushButton_pt_etal_prec_clicked(self):
        """
        Slot documentation goes here.
        """

            #Preparation de l'onglet pour une nouvelle saisie :
        
        pt_etal = int(self.lineedit_pt_etal_n_2.text())
        
        
        if pt_etal > 1 :
            
            #Sauvegarde des donnees à l'ecran avant importation pt precedent
            
            donnees_saisies = traitement_donnees_saisies(self)
            sauvegarde_onglet_saisie(donnees_saisies,self.lineedit_pt_etal_n_2.text() )
            
            pt_etal-=1
            #importation des donnees deja saisies
            onglet = lecture_onglet_saisie(str(pt_etal))
           
            if pt_etal == 1 :
                path="AppData/"+str(1)
                os.remove(path)
            else :
                pass
            
            
            #presentation de l onglet
            
            effacement_onglet_saisie (self)
            
            self.lineedit_pt_etal_n_2.setText(str(pt_etal))
            self.label_pt_etal.setText(str(pt_etal))
                #reaffectation des donnees
            reaffectation_donnees_onglet_saisie(self, onglet)

        else :
            pass
    
            
    @pyqtSlot()
    def on_textEdit_mesures_inst_1_textChanged(self):
        """
        Slot documentation goes here.
        """
        compte_nbr_ligne_textedit_onglet_saisie(self, 1)

    
    @pyqtSlot()
    def on_textEdit_mesures_inst_2_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 2)
    
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_16_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 16)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_9_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 9)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_10_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 10)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_13_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 13)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_8_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 8)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_12_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 12)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_11_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 11)
    
    @pyqtSlot()
    def on_textEdit_mesures_etalon_brutes_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 21)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_3_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 3)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_15_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 15)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_20_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 20)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_14_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 14)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_7_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 7)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_4_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 4)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_6_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 6)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_19_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 19)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_5_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 5)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_17_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 17)
    
    @pyqtSlot()
    def on_textEdit_mesures_inst_18_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 18)
    
    @pyqtSlot()
    def on_textEdit_mesures_etalon_corrigees_textChanged(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        compte_nbr_ligne_textedit_onglet_saisie(self, 22)
            
            
    @pyqtSlot()
    def on_pushButton_pt_etal_suiv_clicked(self):
        """
        Les donnees brutes de l'etalon sont corrigees par rapport à un poly ds la bdd
        Les saisies sont mises sous forme de list .
        L ensemble de l'onglet est sauvé dans un dict
        Ce dict est enregistre en local dans un pickler
        """
      
        try :

            donnees_saisies = traitement_donnees_saisies(self)    
    #        print("traitement {}".format(traitement_donnees_saisies(self)))
    #        print(donnees_saisies)
            
            list_textedit_nom_instruments = [self.textEdit_nom_instrum_1, self.textEdit_nom_instrum_2, self.textEdit_nom_instrum_3, 
                                self.textEdit_nom_instrum_4, self.textEdit_nom_instrum_5, self.textEdit_nom_instrum_6, 
                                self.textEdit_nom_instrum_7, self.textEdit_nom_instrum_8, self.textEdit_nom_instrum_9, 
                                self.textEdit_nom_instrum_10, self.textEdit_nom_instrum_11, self.textEdit_nom_instrum_12, 
                                self.textEdit_nom_instrum_13, self.textEdit_nom_instrum_14, self.textEdit_nom_instrum_15, 
                                self.textEdit_nom_instrum_16, self.textEdit_nom_instrum_17, self.textEdit_nom_instrum_18,
                                self.textEdit_nom_instrum_19, self.textEdit_nom_instrum_20
                                ]
            
            i=0
            while i < self.SpinBox_nbr_instruments.value() :
                
                instrument = {}
                
                instrument["mesures_etal_brute"] =donnees_saisies["mesures_etal_brute"]
                instrument ["mesures_etal_corri"] = donnees_saisies["mesures_etal_corri"]
                instrument ["moyenne_etal_corri"] = donnees_saisies["moyenne_etal_corri"]
                instrument ["moyenne_etal_brute"] = donnees_saisies["moyenne_etal_brute"] 
                instrument ["operateur"] = donnees_saisies["operateur"] 
                instrument ["generateur"] = donnees_saisies["generateur"]
                instrument ["etalon"] = donnees_saisies["etalon"]
                instrument ["pt_etalonnag_n"] = donnees_saisies["pt_etalonnag_n"] 
                instrument ["nb_acquisition"] = donnees_saisies["nb_acquisition"]
                instrument ["temp_consig"] = donnees_saisies["temp_consig"]
                instrument["nom_instrument"+"_"+str(i+1)] = list_textedit_nom_instruments[i].toPlainText()
                instrument ["moyenne_instrum"+"_"+str(i+1)] = donnees_saisies["moyenne_instrum"+"_"+str(i+1)]
                instrument ["mesures_instrum"+"_"+str(i+1)] = donnees_saisies["mesures_inst"+"_"+str(i+1)]               
                
                instrument ["corrections_instrum"+"_"+str(i+1)] = donnees_saisies["corrections_instrum"+"_"+str(i+1)]
                instrument ["moyenne_corrections_instrum"+"_"+str(i+1)] = donnees_saisies["moyenne_corrections_instrum"+"_"+str(i+1)]
                instrument ["ecartype_corrections_instrum"+"_"+str(i+1)] = donnees_saisies["ecartype_corrections_instrum"+"_"+str(i+1)]
                instrument["fuite_thermique_instrum"+"_"+str(i+1)] = donnees_saisies["fuite_thermique_instrum"+"_"+str(i+1)]                                                                        
                instrument["auto_echauffement_instrum"+"_"+str(i+1)]  = donnees_saisies["auto_echauffement_instrum"+"_"+str(i+1)]               
                
                nom_fichier = instrument ["pt_etalonnag_n"] +"_"+"instrum"+"_"+str(i+1)
    
                sauvegarde_onglet_saisie(instrument,nom_fichier ) 
           
                i+=1
                
                
            #Sauvegarde des saisies dans un pickler (un fichier pour longlet et un fichier par instrum etalon avec)
            sauvegarde_onglet_saisie(donnees_saisies,self.lineedit_pt_etal_n_2.text() )
            
            
            #Preparation de l'onglet pour une nouvelle saisie :
            pt_etal = int(self.lineedit_pt_etal_n_2.text())
            nbr_pt = int(self.SpinBox_nbr_pt_etal.value())
            
            if pt_etal < nbr_pt :
            
                pt_etal+=1
                self.lineedit_pt_etal_n_2.setText(str(pt_etal))
                self.label_pt_etal.setText(str(pt_etal))
                            
                #effacement des donnees presentes
                effacement_onglet_saisie (self)
                
                
                #test si un fichier de sauvegarde existe dans ce cas on l'importe et on l'affiche sert pour la navigation dans les onglets apres une saisie
                path ="Appdata/"+str(pt_etal)
                if os.path.isfile(path):
                    
                    onglet = lecture_onglet_saisie(str(pt_etal))
                    reaffectation_donnees_onglet_saisie(self, onglet)
                    os.remove(path)
                    
                
                else  :
                    pass
                
                           
                         
            else :
                
                QMessageBox.information(self, 
                    self.trUtf8("Saisie étalonnage"), 
                    self.trUtf8("Merci la saise est terminée")) 
               
            #on affiche l'onglet resultats
                self.tabWidget.setCurrentIndex(2)
                nom_premier_instrum = list_textedit_nom_instruments[0].toPlainText()                
                self.qlabel_nom_instrument_onglet_resultats.setText(nom_premier_instrum)
                onglet_resultat(self, 1)
                 
        except TypeError :
            QMessageBox.critical(self, 
                                            self.trUtf8("Attention"),
                                            self.trUtf8("Merci de verifier vos saisies"))
            
            pass
           
     
    @pyqtSlot(int)
    def on_SpinBox_nbr_instruments_valueChanged(self, p0):
        """
        Donne le nbr d''instrument à etalonner gere la possibilité ou non d'utiliser les combobox
        """
        list_combobox=[self.comboBox_ident_instrum_1, self.comboBox_ident_instrum_2, self.comboBox_ident_instrum_3, 
                                self.comboBox_ident_instrum_4, self.comboBox_ident_instrum_5, self.comboBox_ident_instrum_6, 
                                self.comboBox_ident_instrum_7, self.comboBox_ident_instrum_8, self.comboBox_ident_instrum_9, 
                                self.comboBox_ident_instrum_10, self.comboBox_ident_instrum_11, self.comboBox_ident_instrum_12, 
                                self.comboBox_ident_instrum_13, self.comboBox_ident_instrum_14, self.comboBox_ident_instrum_15, 
                                self.comboBox_ident_instrum_16, self.comboBox_ident_instrum_17, self.comboBox_ident_instrum_18,
                                self.comboBox_ident_instrum_19, self.comboBox_ident_instrum_20, 
                                ]
        list_groupbox = [ self.groupBox_instrument_1, self.groupBox_instrument_2, self.groupBox_instrument_3, self.groupBox_instrument_4, 
                                self.groupBox_instrument_5, self.groupBox_instrument_6, self.groupBox_instrument_7, self.groupBox_instrument_8, 
                                self.groupBox_instrument_9, self.groupBox_instrument_10, self.groupBox_instrument_11, self.groupBox_instrument_12,
                                self.groupBox_instrument_13, self.groupBox_instrument_14, self.groupBox_instrument_15, self.groupBox_instrument_16,
                                self.groupBox_instrument_17, self.groupBox_instrument_18, self.groupBox_instrument_19, self.groupBox_instrument_20
                                ]
        
        nbr_instrum = self.SpinBox_nbr_instruments.value()
        for elem in list_combobox:
            elem.setEnabled(False)
            
        for ele in list_groupbox :
            ele.setEnabled(False)
            
            
        i=0
        while i <= nbr_instrum-1:
            list_combobox[i].setEnabled(True)
            list_groupbox[i].setEnabled(True)
            i+=1
        
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_19_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 19)
        
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_20_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 20)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_8_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 8)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_11_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 11)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_9_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 9)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_15_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 15)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_7_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 7)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_17_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 17)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_14_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 14)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_16_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 16)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_18_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 18)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_5_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 5)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_12_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 12)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_4_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 4)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_3_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 3)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_10_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 10)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_1_activated(self, p0):
        """
        Slot documentation goes here.
        """
        
        gestion_caracteristiques_instrument(self, 1)

        
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_13_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 13)
    
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_2_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 2)
    
    @pyqtSlot()
    def on_pushButton_inst_prec_clicked(self):
        """
        Slot documentation goes here.
        """
        nom_instrument = [self.comboBox_ident_instrum_1.currentText (), self.comboBox_ident_instrum_2.currentText (), self.comboBox_ident_instrum_3.currentText (), 
                                    self.comboBox_ident_instrum_4.currentText (), self.comboBox_ident_instrum_5.currentText (), self.comboBox_ident_instrum_6.currentText (), 
                                    self.comboBox_ident_instrum_7.currentText (), self.comboBox_ident_instrum_8.currentText (), self.comboBox_ident_instrum_9.currentText (), 
                                    self.comboBox_ident_instrum_10.currentText (), self.comboBox_ident_instrum_11.currentText (), self.comboBox_ident_instrum_12.currentText (),
                                    self.comboBox_ident_instrum_13.currentText () , self.comboBox_ident_instrum_14.currentText (), self.comboBox_ident_instrum_15.currentText (), 
                                    self.comboBox_ident_instrum_16.currentText (), self.comboBox_ident_instrum_17.currentText (), self.comboBox_ident_instrum_18.currentText (), 
                                    self.comboBox_ident_instrum_19.currentText (), self.comboBox_ident_instrum_20.currentText ()
                                    ]
        
        #lecture et sauvegarde de l'onglet resultat deja saisie
        num_instrument = int(self.label_5_inst_n.text())
        sauvegarde_onglet_resultat(self, num_instrument)
        
       
        # Preparation de l'onglet avec l'instrument precedent
        inst_prec = num_instrument - 1
        indice_nom_instrument = inst_prec - 1
        self.qlabel_nom_instrument_onglet_resultats.setText(nom_instrument[indice_nom_instrument])
        
        if inst_prec > 0 :
            #manque la sauvegarde à ce niveau
            nom_fichier = str("resultat_" + "_instrum_"+ str(inst_prec))
            lecture_onglet_saisie(nom_fichier)
            efface_onglet_resultat (self)
            self.label_5_inst_n.setText(str(inst_prec))
            onglet_resultat(self, inst_prec)
        else :
            pass


    @pyqtSlot()
    def on_button_valid_config_clicked(self):
        """
        Apres appuis sur le bouton on créé un fichier de configuration et on passe à l'onglet saisie .
        Sur celui ci nous mettons a jour le nom des instruments.
        """
                
        try:
            
            list_combobox_configuration = [self.comboBox_ident_instrum_1, self.comboBox_ident_instrum_2, self.comboBox_ident_instrum_3, 
                                                        self.comboBox_ident_instrum_4, self.comboBox_ident_instrum_5, self.comboBox_ident_instrum_6, 
                                                        self.comboBox_ident_instrum_7, self.comboBox_ident_instrum_8, self.comboBox_ident_instrum_9, 
                                                        self.comboBox_ident_instrum_10, self.comboBox_ident_instrum_11, self.comboBox_ident_instrum_12, 
                                                        self.comboBox_ident_instrum_13, self.comboBox_ident_instrum_14, self.comboBox_ident_instrum_15, 
                                                        self.comboBox_ident_instrum_16, self.comboBox_ident_instrum_17, self.comboBox_ident_instrum_18, 
                                                        self.comboBox_ident_instrum_19, self.comboBox_ident_instrum_20]
                
            #test si un instrument a ete choisi plusieurs fois  si oui on leve une exception Value Error
            list_instrument = []            
            i=0
            while i < self.SpinBox_nbr_instruments.value():
                list_instrument.append(list_combobox_configuration[i].currentText())
                i+=1
            set_instrument = set(list_instrument)
            
            if len(set_instrument) != len(list_instrument):
                raise ValueError
                
            #on sauvegarde l'onglet configuration et on recuperes ses donnees dans une list de dictionnaire
            donnees = lecture_onglet_configuration(self) 
            
            #Acces bdd
            db = GestionBdd('db')
            db.reconnexion()
            db.modif_table_instruments(donnees)
            sauvegarde_onglet_saisie(donnees , "configuration")
            
                #remarque la resolution des instruments est contenue dans donnees[i+1]["resolution"]
            
            i = 0
            erreur = []
            compteur = 0
            while i < donnees[0]["nbr_instrum"]:
                valeur_textedit_resolution = []
                immersion = []
    
                for ele in donnees[i+1]["resolution"]:
                    valeur_textedit_resolution.append(ele)
                 
                if len(valeur_textedit_resolution) < 1 :
                    erreur.append(True)
                    compteur +=1
                else :
                    erreur.append(False)
                               
                for ele in donnees[i+1]["immersion"]:
                    immersion.append(ele)
                               
                if len(immersion) == 0: # on verifie qu'il y est des donnees saisies dans le textbox sinon erreur
                    raise IndexError
                   
                i+=1
    
            
            if compteur > 0 :                
                QMessageBox.critical(self, 
                                                self.trUtf8("Attention"),
                                                self.trUtf8("Merci de verifier les valeurs des resolutions saisies "))
            else :                
                self.tabWidget.setCurrentIndex(1)                   
                self.tab_2.setEnabled(True)
                self.lineedit_pt_etal_n_2.setText("1")
                self.label_pt_etal.setText("1")
                nbr_pt_temp = self.SpinBox_nbr_pt_etal.value()
                self.label_nb_pt_etal.setText(str(nbr_pt_temp))
                
                list_textedit_mesures_instrum = [self.textEdit_mesures_inst_1, self.textEdit_mesures_inst_2,  self.textEdit_mesures_inst_3, 
                                        self.textEdit_mesures_inst_4,  self.textEdit_mesures_inst_5,  self.textEdit_mesures_inst_6, 
                                        self.textEdit_mesures_inst_7,  self.textEdit_mesures_inst_8,  self.textEdit_mesures_inst_9, 
                                        self.textEdit_mesures_inst_10,  self.textEdit_mesures_inst_11,  self.textEdit_mesures_inst_12, 
                                        self.textEdit_mesures_inst_13,  self.textEdit_mesures_inst_14,  self.textEdit_mesures_inst_15, 
                                        self.textEdit_mesures_inst_16,  self.textEdit_mesures_inst_17,  self.textEdit_mesures_inst_18, 
                                        self.textEdit_mesures_inst_19,  self.textEdit_mesures_inst_20 ]
                                        
    #Remarque : Les noms des instruments sont contenus dans donnees[i+1]["nom_instrum"]
                
    
                list_textedit_nom_instrum = [self.textEdit_nom_instrum_1,self.textEdit_nom_instrum_2, self.textEdit_nom_instrum_3, 
                                                        self.textEdit_nom_instrum_4, self.textEdit_nom_instrum_5, self.textEdit_nom_instrum_6, 
                                                        self.textEdit_nom_instrum_7, self.textEdit_nom_instrum_8, self.textEdit_nom_instrum_9,
                                                        self.textEdit_nom_instrum_10, self.textEdit_nom_instrum_11, self.textEdit_nom_instrum_12, 
                                                        self.textEdit_nom_instrum_13, self.textEdit_nom_instrum_14, self.textEdit_nom_instrum_15, 
                                                        self.textEdit_nom_instrum_16, self.textEdit_nom_instrum_17, self.textEdit_nom_instrum_18, 
                                                        self.textEdit_nom_instrum_19, self.textEdit_nom_instrum_20]
                                                        
                list_line_textedit_nom_instrum_onglet_conformite = [self.lineEdit_instrum_1, self.lineEdit_instrum_2, self.lineEdit_instrum_3,
                                                                                        self.lineEdit_instrum_4, self.lineEdit_instrum_5, self.lineEdit_instrum_6, 
                                                                                        self.lineEdit_instrum_7, self.lineEdit_instrum_8, self.lineEdit_instrum_9, 
                                                                                        self.lineEdit_instrum_10, self.lineEdit_instrum_11, self.lineEdit_instrum_12, 
                                                                                        self.lineEdit_instrum_13, self.lineEdit_instrum_14, self.lineEdit_instrum_15, 
                                                                                        self.lineEdit_instrum_16, self.lineEdit_instrum_17, self.lineEdit_instrum_18, 
                                                                                        self.lineEdit_instrum_19, self.lineEdit_instrum_20]
                
                
                nbr_instrum = donnees[0]["nbr_instrum"]
                
                #compteur i represente l'indice de la list contenant les differents textedit
                i = 0
                while i <= 19:
                    list_textedit_mesures_instrum[i].setEnabled(True)
                    i +=1
                                   
                i = self.SpinBox_nbr_instruments.value()  #on va reregarder dans une list indice par a 0
                while i <= 19 :
                    list_textedit_mesures_instrum[i].setEnabled(False)
                    i +=1
                
                #compteur pour affectation des noms des instruments choisis dans longlet saisie et conformité
                j=0
                while j < nbr_instrum :
                    list_textedit_nom_instrum[j].clear()
                    list_line_textedit_nom_instrum_onglet_conformite[j].clear()
                    
                    #Remarque : Les noms des instruments sont contenus dans donnees[i+1]["nom_instrum"]
                    list_textedit_nom_instrum[j].append(donnees[j+1]["nom_instrum"])
                    list_line_textedit_nom_instrum_onglet_conformite[j].setText(donnees[j+1]["nom_instrum"])
                    
                    #on nettoie les textedit mesure instrum /et etalon
                    list_textedit_mesures_instrum[j].clear()
                    self.textEdit_mesures_etalon_brutes.clear()
                    self.textEdit_mesures_etalon_corrigees.clear()
                    
                    j+=1
                
                
                #verification si un fichier avec des mesures existe
                path =os.path.abspath("AppData/1") #permet de recuperer le chemin absolu du fichier le test ne marche pas sinon
    
                if os.path.isfile(path) ==True :                                           
                    onglet = lecture_onglet_saisie(str(1))
                    reaffectation_donnees_onglet_saisie(self, onglet)  
                else :
                    pass
                    
                if self.type_de_saisie[0] == 3: #s'il s'agit d'une moficationd es donnees apres une validation (un seul instrument)
                    self.Zone_text_temp_consigne_2.setEnabled(False)
                    self.Combobox_operateur_select_2.setEnabled(False)
                    self.Combobox_generateur_select_2.setEnabled(False)
                    self.Combobox_etalon_select_2.setEnabled(False)
                    self.textEdit_mesures_etalon_brutes.setEnabled(False)
        except IndexError:
            QMessageBox.critical(self, 
                                                self.trUtf8("Attention"),
                                                self.trUtf8("Merci de verifier les saisies d'immersion "))
        except ValueError:
            QMessageBox.critical(self, 
                                                self.trUtf8("Attention"),
                                                self.trUtf8("Plusieurs instruments identiques ont ete selectionnés"))
            
            
    @pyqtSlot(str)
    def on_comboBox_ident_instrum_6_activated(self, p0):
        """
        Slot documentation goes here.
        """
        gestion_caracteristiques_instrument(self, 6)
        
    

    @pyqtSlot()
    def on_textEdit_fuite_thermique_selectionChanged(self):
        """
        Slot documentation goes here.
        """
        valeur_textEdit_corrections = self.textEdit_corrections.toPlainText()
 
        self.textEdit_fuite_thermique.clear()
        tuple_1 = QInputDialog.getItem(self, 
                       self.trUtf8("Premiere Valeur"), 
                       self.trUtf8("Choisissez dans la liste"),
                       valeur_textEdit_corrections.split())
        tuple_2 = QInputDialog.getItem(self, 
                       self.trUtf8("Deuxieme Valeur"), 
                       self.trUtf8("Choisissez dans la liste"),
                       valeur_textEdit_corrections.split())
                       
        valeur_1 = tuple_1[0]
        valeur_2 = tuple_2[0]
        delta = float(valeur_2) - float(valeur_1)
        fuite_thermique = str(round((numpy.abs(delta)/(numpy.sqrt(3))), 12))
        
        i=0
        while i < self.SpinBox_nbr_pt_etal.value() :
            self.textEdit_fuite_thermique.append(fuite_thermique)
            i+=1
        nvlle_incertitude =  calcul_incertitude(self)
            #incertitude
        self.textEdit_U.clear()
        for ele in nvlle_incertitude :           
            self.textEdit_U.append(str(round(ele, 12)))
    
    @pyqtSlot()
    def on_pushButton_inst_suiv_clicked(self):
        """
        Slot documentation goes here.
        """
        nom_instrument = [self.comboBox_ident_instrum_1.currentText (), self.comboBox_ident_instrum_2.currentText (), self.comboBox_ident_instrum_3.currentText (), 
                                    self.comboBox_ident_instrum_4.currentText (), self.comboBox_ident_instrum_5.currentText (), self.comboBox_ident_instrum_6.currentText (), 
                                    self.comboBox_ident_instrum_7.currentText (), self.comboBox_ident_instrum_8.currentText (), self.comboBox_ident_instrum_9.currentText (), 
                                    self.comboBox_ident_instrum_10.currentText (), self.comboBox_ident_instrum_11.currentText (), self.comboBox_ident_instrum_12.currentText (),
                                    self.comboBox_ident_instrum_13.currentText () , self.comboBox_ident_instrum_14.currentText (), self.comboBox_ident_instrum_15.currentText (), 
                                    self.comboBox_ident_instrum_16.currentText (), self.comboBox_ident_instrum_17.currentText (), self.comboBox_ident_instrum_18.currentText (), 
                                    self.comboBox_ident_instrum_19.currentText (), self.comboBox_ident_instrum_20.currentText ()
                                    ]
        
        #lecture et sauvegarde de l'onglet resultat deja saisie
        n_instrument = int(self.label_5_inst_n.text())
        sauvegarde_onglet_resultat(self, n_instrument)
        
        # Preparation de l'onglet avec l'instrument suivant
        inst_suiv = n_instrument +1        
        
        if inst_suiv <= self.SpinBox_nbr_instruments.value() :
            efface_onglet_resultat (self)
            self.label_5_inst_n.setText(str(inst_suiv))
            onglet_resultat(self, inst_suiv)
            indice_nom_instrument = inst_suiv - 1 #indice de la list par à 0
            self.qlabel_nom_instrument_onglet_resultats.setText(nom_instrument[indice_nom_instrument])
        else :
            if self.type_de_saisie[0] == 1: #saisie normale
                self.tabWidget.setCurrentIndex(3)
                self.actionEnregistrer.setEnabled(True)
            
            elif self.type_de_saisie[0] == 2: #modif par n° campagne
                self.tabWidget.setCurrentIndex(3)
#                self.actionEnregistrer.setEnabled(False)
                self.actionMise_jour.setEnabled(True)
                
            elif self.type_de_saisie[0] == 3: # modif par certficiat
                self.tabWidget.setCurrentIndex(3)
#                self.actionEnregistrer.setEnabled(False)
                self.actionMise_jour.setEnabled(True) 
                
            elif self.type_de_saisie[0] == 4:
                self.tabWidget.setCurrentIndex(3)
#                self.actionEnregistrer.setEnabled(False)
                self.actionAnnule_et_Remplace.setEnabled(True)
    
    @pyqtSlot()
    def on_actionMise_jour_triggered(self):
        """
        Slot documentation goes here.
        """
        #Selection du dossier pour la sauvegarde des certificats:
        
        dossier = QFileDialog.getExistingDirectory(None ,  "Selectionner le dossier de sauvegarde des Rapports", 'y:/1.METROLOGIE/0.ARCHIVES ETALONNAGE-VERIFICATIONS/1-TEMPERATURE/')
        if dossier !="":
            #gestion progressbar
            self.progressBar.setMaximum(self.SpinBox_nbr_instruments.value())
            self.progressBar.setMinimum(0)
            self.progressBar.setValue(0)
            self.tabWidget.setCurrentIndex(4)
            
            #Acces bdd
            db = GestionBdd('db')
            db.reconnexion()
            #constantes reutilisables
            date = self.calendarWidget.selectedDate()
            date_etal = date.toString('yyyy-MM-dd')
            date_etal_fr = date.toString('dd/MM/yyyy')
    #        id_campagne_etal = self.type_de_saisie[2] #les listes partent avec un indice 0
            nom_fichier = "configuration"
            config_etal = lecture_onglet_saisie(nom_fichier)
    
            
                
            if self.type_de_saisie[0] == 2:            
              
               #
               
               #Mise à jour Table  CAMPAGNE_ETALONNAGE_TEMP
                nom_campagne = self.type_de_saisie[1]
                id_campagne_etal = self.type_de_saisie[2]
                
                #recuperation des id des instruments avant modification
                list_id_instrum_avant_modif = db.recuperation_id_instrums_table_campagne_etalonnage_temp(id_campagne_etal)

                #Gestion onglet conclusion :
                self.qlabel_campagne_etal_onglet_conclusion.setText(nom_campagne)
                  
                
                #gestion des id instrum apres modifications
                list_instrum = []
                i = 0
                while i < self.SpinBox_nbr_instruments.value():
                    list_instrum.append(str(config_etal[i+1]["nom_instrum"]))                
                    i+=1
                    
                    # on recherche les ID_instrument            
                id_instrum = []
                for ele in list_instrum:
                    id = db.recherche_id_instrument(ele)
                    id_instrum.append(id)    
                
                compteur = 20 - self.SpinBox_nbr_instruments.value()
                       
                i= 0
                while i < compteur: 
                    id_instrum.append(0)
                    i+=1
                db.mise_a_jour_table_campagne_etalonnage_temp(nom_campagne, date_etal, id_instrum, id_campagne_etal)
                
                #####################################
                #Mise à jour table Etalonnage Temp Administration :
                #recuperation N de certificat :
                
                
                i = 0            
                while i < self.SpinBox_nbr_instruments.value() :                       
                    
                    #progress bar
                    self.progressBar.setValue(i)
                    
                    #preparation CE/CV
                    rapport_etalonnage = {} #dictionnaire ou on met les donnees afin de remplir le CE/CV
                    #Lists pour remplir le CE/CV
                    etalon = []
                    generateur = []
                    valeurs_etalon_corri = []
                    valeurs_instrum = []
                    valeurs_corrections = []
                    valeurs_U = []
                    valeurs_resolution = []
                    
                    nom_instrument_avant_modif = db.recherche_identification_instrument(list_id_instrum_avant_modif[i])#permet de s'affranchir du probleme lors modification identifiant instrument
                    
                    nom_fichier = str("resultat_" + "_instrum_"+ str(i+1)) #resultat__instrum_1
                    donnees = lecture_onglet_saisie(nom_fichier)
                    
                    
                    #champs de la base :
                    nbr_pt_etalonnage = str("'"+str(self.SpinBox_nbr_pt_etal.value())+"'")           
                    identification_instrum = str("'"+str(donnees[0]["nom_instrument_"+str(i+1)])+"'")
                    rapport_etalonnage["identification_instrument"] = str(donnees[0]["nom_instrument_"+str(i+1)])
                    rapport_etalonnage["renseignement_complementaire"] = config_etal[i+1]["renseignement_complementaire"]
                    rapport_etalonnage["Etat_reception"] = config_etal[i+1]["Etat_reception"]
                    #requete pour avoir le n°serie , le constructeur ,designation , type , code du client:
                    caracteristique_instrum = db.recherche_caract_instrument(identification_instrum)
                    etat_reception = (config_etal[i+1]["Etat_reception"].replace('"', '\"')).replace('\'','\'\'')
                    n_serie =  caracteristique_instrum[0]
                    rapport_etalonnage["n_serie"] = caracteristique_instrum[1]
                    rapport_etalonnage["constructeur"] = caracteristique_instrum[2]
                    rapport_etalonnage["designation"] = caracteristique_instrum[3]
                    rapport_etalonnage["type"] = caracteristique_instrum[4]
                    rapport_etalonnage["affectation"] = caracteristique_instrum[5]
                    code_client = caracteristique_instrum[6]
                   
                    #technicien
                    id_technicien = str(donnees[0]['operateur']+1)
                    
                    caract_technicien = db.recherche_technicien_depuis_id(id_technicien)
                   
                    technicien =  caract_technicien[0]
                    rapport_etalonnage["operateur"] = caract_technicien[1]
                    
                    #Recherche du client             
                    caract_client = db.recherche_client(code_client)
                    
                    rapport_etalonnage["societe"] = caract_client[0]
                    rapport_etalonnage["adresse"] = caract_client[1]
                    rapport_etalonnage["code_postal"] = caract_client[2]
                    rapport_etalonnage["ville"] = caract_client[3] 
                    
                    
                    #recherche n°inventaire (SAP)
                    n_sap = db.recherche_n_inventaire(identification_instrum, n_serie)
                    if isinstance(n_sap, str):
                        rapport_etalonnage["n_sap"] = n_sap
                    else:
                        n_sap = "Neant"
                        rapport_etalonnage["n_sap"] = n_sap
                        
                    rapport_etalonnage["date_etalonnage"] = date_etal_fr
                    
                     
                    compteur = 0
                    list_pt_etal = []
                    while compteur < self.SpinBox_nbr_pt_etal.value():
                        list_pt_etal.append(donnees[compteur]["temp_consig"])
                        
                        if donnees[compteur]["generateur"] == 3 or donnees[compteur]["generateur"]== 4:
                            nom_procedure = str("'"+"PDL/PIL/SUR/MET/MO/030"+"'")
                            rapport_etalonnage["milieu"] = "AIR" 
                            rapport_etalonnage["n_mode_operatoire"] = "030"
                        else:
                            nom_procedure = str("'"+"PDL/PIL/SUR/MET/MO/025"+"'")
                            rapport_etalonnage["milieu"] = "LIQUIDE" 
                            rapport_etalonnage["n_mode_operatoire"] = "025"
                            
                        compteur +=1
                    
                    rapport_etalonnage["temps_consignes"] =list_pt_etal
                    
                        #on va chercher le N° du certificat avec la fonction numero_certificat
                    
                    num_doc = db.recherche_n_ce_id_campagne_identification_instrum(id_campagne_etal, nom_instrument_avant_modif) #renvoi un tupple avec le n°doc et le numero annule et remplace
                    
                    if num_doc[1] == "Neant":
                        rapport_etalonnage["n_certificat"] = num_doc[0]
                        nom_fichier_ce =  num_doc[0]
                        num_cv =  num_doc[0] + "V" 
                        nom_fichier_cv = num_doc[0] + "V" 
                    else:
                        rapport_etalonnage["n_certificat"] = num_doc[0] + "\n Annule et Remplace le document "+str(num_doc[1])
                        nom_fichier_ce =  str(num_doc[0]) + " Annule et Remplace le document "+str(num_doc[1])
                        num_cv =  str(num_doc[0]) + "V" + "\n Annule et Remplace le document "+str(num_doc[1])+"V"
                        nom_fichier_cv = str(num_doc[0]) + "V" + " Annule et Remplace le document "+str(num_doc[1])+"V"
                    
                    phrase_qlabel = "Le certificat n° {} de l'instrument {} est en cours de création".format(rapport_etalonnage["n_certificat"], rapport_etalonnage["identification_instrument"] )
                    self.qlabel_information_onglet_conclusion.setText(phrase_qlabel)
                    
                    
                   
                        #Test à ecrire savoir si cofrac ou non
                    bouton_radio = [self.radioButton_cofrac_1, self.radioButton_cofrac_2, self.radioButton_cofrac_3, self.radioButton_cofrac_4, 
                                            self.radioButton_cofrac_5, self.radioButton_cofrac_6, self.radioButton_cofrac_7, self.radioButton_cofrac_8, 
                                            self.radioButton_cofrac_9, self.radioButton_cofrac_10, self.radioButton_cofrac_11, self.radioButton_cofrac_12, 
                                            self.radioButton_cofrac_13, self.radioButton_cofrac_14, self.radioButton_cofrac_15, self.radioButton_cofrac_16, 
                                            self.radioButton_cofrac_17, self.radioButton_cofrac_18, self.radioButton_cofrac_19, self.radioButton_cofrac_20]
                    
                    if bouton_radio[i].isChecked() :
                        edit_ce = str("'"+"COFRAC"+"'")
                        edit_cv =str("'"+"COFRAC"+"'")
                    else :
                        edit_ce = str("'"+"NON_COFRAC"+"'")
                        edit_cv =str("'"+"NON_COFRAC"+"'")
        
                    
                    ville_etal = str("'"+"Le Mans"+"'")
                    domaine_accreditation = str("'"+"Temperature"+"'")
                    if edit_ce == str("'"+"COFRAC"+"'"):
                        numero_accreditation = str("'"+"2-1950"+"'")
                    else : 
                        numero_accreditation = str("'"+"Neant"+"'")
                    nbr_pages_annexe = str("'"+str(0)+"'")
                    laboratoire = str("'"+"Laboratoire de Metrologie"+"'")
                    
                    #Requete pour mise à jour des donnees dans la table etalonnage_temp_administration
                
                    #Recuperation de l'id_etalonnage_temp_administration :
                    id_etalonnage_temp_administration = db.recherche_id_etalonnage_admin_id_campagne_identification_instrum(id_campagne_etal, nom_instrument_avant_modif)
                    rapport_etalonnage["id_etal"] = id_etalonnage_temp_administration
                    
                    
                    db.mise_a_jour_table_etalonnage_temp_administration(id_campagne_etal, nom_instrument_avant_modif, nbr_pt_etalonnage, 
                                                                                        n_sap, identification_instrum, date_etal, 
                                                                                        nom_procedure, num_doc, edit_ce, edit_cv, 
                                                                                        technicien, ville_etal, domaine_accreditation, 
                                                                                        numero_accreditation, nbr_pages_annexe, 
                                                                                        laboratoire, etat_reception)
                ###########################################################
                    #Mise à jour Table ETALONNAGE_MESURE
                            #champs : CORRECTION, NUM_DOCUMENT, IDENTIFICATION, NBR_PT_ETALONNAGE, NUM_PT_ETAL, NBR_MESURE_PAR_PT, NUM_MESURE, MESURE_ETAL_NC, MESURE_ETAL_C, MESURE_INSTRUM
                
                    
                
                    #gestion profondeur immersion            
                    immersion = []
                    qlabel_prof_immersion = [self.textEdit_prof_imm_instrum_1, self.textEdit_prof_imm_instrum_2, self.textEdit_prof_imm_instrum_3, 
                                                    self.textEdit_prof_imm_instrum_4, self.textEdit_prof_imm_instrum_5, self.textEdit_prof_imm_instrum_6, 
                                                    self.textEdit_prof_imm_instrum_7, self.textEdit_prof_imm_instrum_8, self.textEdit_prof_imm_instrum_9, 
                                                    self.textEdit_prof_imm_instrum_10, self.textEdit_prof_imm_instrum_11, self.textEdit_prof_imm_instrum_12, 
                                                    self.textEdit_prof_imm_instrum_13, self.textEdit_prof_imm_instrum_14, self.textEdit_prof_imm_instrum_15, 
                                                    self.textEdit_prof_imm_instrum_16, self.textEdit_prof_imm_instrum_17, self.textEdit_prof_imm_instrum_18, 
                                                    self.textEdit_prof_imm_instrum_19, self.textEdit_prof_imm_instrum_20 ]
                
                    for ele in qlabel_prof_immersion[i].toPlainText().split() : 
                        immersion.append(ele)            
                
    
                    #gestion resolutions instrument saisie
                    resolution_instrument = []
                    qlabel_resolution_instrum = [self.textEdit_resolution_instrum_1, self.textEdit_resolution_instrum_2, self.textEdit_resolution_instrum_3, 
                                                        self.textEdit_resolution_instrum_4, self.textEdit_resolution_instrum_5, self.textEdit_resolution_instrum_6, 
                                                        self.textEdit_resolution_instrum_7, self.textEdit_resolution_instrum_8, self.textEdit_resolution_instrum_9, 
                                                        self.textEdit_resolution_instrum_10, self.textEdit_resolution_instrum_11, self.textEdit_resolution_instrum_12, 
                                                        self.textEdit_resolution_instrum_13, self.textEdit_resolution_instrum_14, self.textEdit_resolution_instrum_15, 
                                                        self.textEdit_resolution_instrum_16, self.textEdit_resolution_instrum_17, self.textEdit_resolution_instrum_18, 
                                                        self.textEdit_resolution_instrum_19, self.textEdit_resolution_instrum_20 ]
                    
                    for ele in qlabel_resolution_instrum[i].toPlainText().split() : 
                        resolution_instrument.append(ele.replace(",", "."))  
                
                
                
                    j=0
                    while j < self.SpinBox_nbr_pt_etal.value() :
                        nbr_de_mesure = donnees[j]["nb_acquisition"]
                    
                        k = 0
                        while k < nbr_de_mesure :
                        
                            correction = "'"+str(donnees[j]["corrections_instrum_"+str(i+1)][k])+"'"
                            num_pt_etalonnage ="'"+str(donnees[j]["pt_etalonnag_n"])+"'"
                            nbr_mesure = "'"+str(donnees[j]["nb_acquisition"])+"'"
                            numero_mesure = "'"+str(k+1)+"'"
                            mesure_etal_non_corri = "'"+str(donnees[j]["mesures_etal_brute"][k])+"'"
                            mesure_etal_corri =  "'"+str(donnees[j]["mesures_etal_corri"][k])+"'"
                            mesure_instrum = "'"+str(donnees[j]["mesures_instrum_"+str(i+1)][k])+"'"
    
                            #Mise à jour des donnees de mesures dans la table ETALONNAGE_MESURE 
                        
                            #Recherche de l'id ETAL MESURES
                            id_etal_mesures = db.recherche_id_etalonnage_mesure_id_campagne_id_etal(id_campagne_etal, id_etalonnage_temp_administration, numero_mesure, num_pt_etalonnage)
                            
                            db.mise_a_jour_table_etalonnage_mesure(id_etalonnage_temp_administration, id_campagne_etal, 
                                                                                correction, num_doc, identification_instrum, nbr_pt_etalonnage,
                                                                                num_pt_etalonnage, nbr_mesure, numero_mesure, mesure_etal_non_corri,
                                                                                mesure_etal_corri, mesure_instrum, id_etal_mesures)
    
                                        
                            k += 1
                    ###############################################################
                    #Mise à jour Table ETALONNAGE_RESULTAT
                    
                                #Remplissage Table etalonnage resultats:
                    
                       #champs
                                    
                        nom_etalon = "'"+str(self.Combobox_etalon_select_2.itemText(donnees[j]["etalon"]))+"'"
                        etalon.append(str(self.Combobox_etalon_select_2.itemText(donnees[j]["etalon"])))
                    
                        nom_generateur = "'"+str(self.Combobox_generateur_select_2.itemText(donnees[j]["generateur"]))+"'"
                        generateur.append(str(self.Combobox_generateur_select_2.itemText(donnees[j]["generateur"])))
                    
                        autoechauffement = "'"+str(donnees[j]["auto_echauffement"])+"'"
                    
                        fuite_thermique ="'"+str(donnees[j]["fuite_thermique"])+"'" 
                    
                        resolution = "'"+str(donnees[j]["Resolution"])+"'" # Attention il sagit de l'incertitude et non de la reso de l 'appareil
                        valeurs_resolution.append(donnees[j]["Resolution"])
                    
                        stabilite = "'"+str(donnees[j]["STAB"])+"'"
                    
                        eiv = "'"+str(donnees[j]["EIV"])+"'"
                    
                        moyenne_etalon_non_corri = "'"+str(donnees[j]["moyenne_etal_brute"])+"'"
                                    
                        moyenne_etalon_corri =  "'"+str(donnees[j]["moyenne_etal_corri"])+"'"
                        valeurs_etalon_corri.append(donnees[j]["moyenne_etal_corri"])
                    
                        moyenne_instrum = "'"+str(donnees[j]["moyenne_instrum_"+str(i+1)])+"'"
                        valeurs_instrum.append(donnees[j]["moyenne_instrum_"+str(i+1)])
                    
                        moyenne_correction = "'"+str(donnees[j]["moyenne_corrections_instrum_"+str(i+1)])+"'"
                        valeurs_corrections.append(donnees[j]["moyenne_corrections_instrum_"+str(i+1)])
                    
                        U = "'"+str(donnees[j]["U"])+"'"
                        valeurs_U.append(donnees[j]["U"])
                    
                        ecart_type = "'"+str(donnees[j]["ecart_type"])+"'"
                    
                        temp_consigne = donnees[j]["temp_consig"]
                    
                        valeur_immersion = immersion[j]
                    
                        valeur_resolution_instrument = resolution_instrument[j]
    
                        #insertion dans la table:
                    
                        id_etal_result = db.recherche_id_etal_result_table_etalonnage_resultat(id_campagne_etal, id_etalonnage_temp_administration)
                        db.mis_a_jour_table_etalonnage_resultat(temp_consigne, id_campagne_etal, id_etalonnage_temp_administration,
                                                                        nom_etalon, nom_generateur, autoechauffement , fuite_thermique, 
                                                                        resolution, stabilite, eiv, num_doc, identification_instrum, date_etal, 
                                                                        nbr_pt_etalonnage, num_pt_etalonnage, moyenne_etalon_non_corri, 
                                                                        moyenne_etalon_corri, moyenne_instrum, moyenne_correction, U, 
                                                                        ecart_type, valeur_immersion, valeur_resolution_instrument, id_etal_result[j])
    
                        j += 1
                
    #            resolution = lecture_onglet_configuration(self)            
                    rapport_etalonnage["resolution"] =  resolution_instrument[0] # resolution instrument
                
                    rapport_etalonnage["nbr_pt_etalonnage"] = self.SpinBox_nbr_pt_etal.value()
                    rapport_etalonnage["etalon"] = list(set(etalon)) # set : on crée un set qui permet d'eliminer les doublons
                    rapport_etalonnage["generateur"] = list(set(generateur)) # set : on crée un set qui permet d'eliminer les doublons
                    rapport_etalonnage["moyenne_etalon_corri"] = valeurs_etalon_corri
                    rapport_etalonnage["moyenne_instrum"] = valeurs_instrum
                    rapport_etalonnage["moyenne_correction"] =valeurs_corrections
                    rapport_etalonnage["U"] = valeurs_U
                    rapport_etalonnage["immersion"] = immersion
                
                    #Gestion du polynome correction =f(Tlue) : et enregistrement dans BDD            
                    if self.SpinBox_nbr_pt_etal.value() < 5:
                        ordre_poly = 1
                        poly = calcul_polynome (valeurs_instrum,valeurs_corrections, ordre_poly)
                        a = poly[0]
                        b = poly[1]
                        c = 0
                    else:
                        ordre_poly = 2
                        poly = calcul_polynome (valeurs_instrum,valeurs_corrections, ordre_poly)
                        a = poly[0]
                        b = poly[1]
                        c = poly[2]
                
                    date_du_jour = datetime.datetime.today()
                    date_du_jour_mise_en_forme = "'{}-{}-{}'".format(date_du_jour.year, date_du_jour.month, date_du_jour.day)
                    archivage = "'N'"
                
                    #mise à jour la table polynome
                    if caracteristique_instrum[3] == "Etalon":
                        reponse = QMessageBox.question(self, 
                                    self.trUtf8("Information"), 
                                    self.trUtf8("Voulez-vous mettre à jour le polynome de correction"), 
                                    QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
                                    
                        if reponse == QtGui.QMessageBox.Yes:
                            db.mise_a_jour_table_polynome(identification_instrum,date_du_jour_mise_en_forme, num_doc,
                                                        date_etal, ordre_poly,a, b, c, archivage)
                        else:
                            pass
                        
                                     
                    else:
                        db.mise_a_jour_table_polynome(identification_instrum,date_du_jour_mise_en_forme, num_doc,
                                                        date_etal, ordre_poly,a, b, c, archivage)
                
                    
                    list_textedit_conformite =  [self.textEdit_resultat_instru_1, self.textEdit_resultat_instru_2, self.textEdit_resultat_instru_3,
                                                        self.textEdit_resultat_instru_4, self.textEdit_resultat_instru_5, self.textEdit_resultat_instru_6, 
                                                        self.textEdit_resultat_instru_7, self.textEdit_resultat_instru_8, self.textEdit_resultat_instru_9, 
                                                        self.textEdit_resultat_instru_10, self.textEdit_resultat_instru_11, self.textEdit_resultat_instru_12, 
                                                        self.textEdit_resultat_instru_13, self.textEdit_resultat_instru_14, self.textEdit_resultat_instru_15, 
                                                        self.textEdit_resultat_instru_16, self.textEdit_resultat_instru_17, self.textEdit_resultat_instru_18, 
                                                        self.textEdit_resultat_instru_19, self.textEdit_resultat_instru_20]
                                                    
                                                
                    list_combobox_emt = [self.comboBox_ref_emt_1, self.comboBox_ref_emt_2, self.comboBox_ref_emt_3, 
                                            self.comboBox_ref_emt_4, self.comboBox_ref_emt_5, self.comboBox_ref_emt_6, 
                                            self.comboBox_ref_emt_7, self.comboBox_ref_emt_8, self.comboBox_ref_emt_9, 
                                            self.comboBox_ref_emt_10, self.comboBox_ref_emt_11, self.comboBox_ref_emt_12, 
                                            self.comboBox_ref_emt_13, self.comboBox_ref_emt_14, self.comboBox_ref_emt_15, 
                                            self.comboBox_ref_emt_16, self.comboBox_ref_emt_17, self.comboBox_ref_emt_18, 
                                            self.comboBox_ref_emt_19, self.comboBox_ref_emt_20]
                    #gestion textedit conformite:
                    resultat_conformite = []
                    for ele in list_textedit_conformite[i].toPlainText().split():                
                        resultat_conformite.append(ele)            
                    
                    if not resultat_conformite:
                     rapport_etalonnage["conformite"] = "Néant"
                    
                    else:
                        if resultat_conformite[0] == "Non":
                            rapport_etalonnage["conformite"] = resultat_conformite[0] + " " + resultat_conformite[1]
                        else:
                            rapport_etalonnage["conformite"] = resultat_conformite[0]
                    
                
                    #Recuperation nom referentiel            
                    nom_referentiel_conformite = list_combobox_emt[i].currentText()
                    id_referentiel_conformite = db.recherche_id_ref_emt(nom_referentiel_conformite)
                    rapport_etalonnage["referentiel_emt"] = db.recherche_nom_ref_emt(id_referentiel_conformite)
               
                    #gestion des CE/CV et la sauvegarde 
                    certificat = RapportEtalonnage(edit_ce) #appel à la classe RapportEtalonnage
                    certificat.mise_en_forme_ce(rapport_etalonnage, dossier, nom_fichier_ce)
    
                    #CV :
                    edition_cv = gestion_edition_cv(self) #retourne une list avec true ou false en fct s'il faut ou non edité le cv
                    #on regarde ds la bdd s'il existe une declaration de conf si oui on a un tupple avec id_conformite, id_ref, conclusion
                    conformite_instrum = db.recherche_id_conformite_temp_resultat_id_etalonnage(id_etalonnage_temp_administration)
    #                print(conformite_instrum)
                    
                    if edition_cv[i] == True:
                        if conformite_instrum[0] == None:
                            db.insertion_table_conformite_temp_resultat(rapport_etalonnage, date_etal)
                        else:
                            db.mise_a_jour_table_conformite_temp_resultat(conformite_instrum [0], rapport_etalonnage, date_etal)
                        constat = RapportEtalonnage(edit_ce)
                        
                        if num_doc[1] == "Neant":
                            constat.mise_en_forme_cv(rapport_etalonnage, dossier, nom_fichier_cv)
                        else:
                            constat.mise_en_forme_cv_annule_remplace(rapport_etalonnage, dossier, num_cv, nom_fichier_cv)
                        constat.mise_en_forme_cv(rapport_etalonnage, dossier, nom_fichier_cv)           
                    
                    i += 1
            
            elif self.type_de_saisie[0] == 3:
                #modification pour un seul certificat
                
                
                identification_ancien_instrum = db.recherche_identification_instrument_num_ce(self.type_de_saisie[1])
                id_ancien_instrum = db.recherche_id_instrument(identification_ancien_instrum)
                
                nom_campagne = self.type_de_saisie[2]
                id_campagne = db.recherche_id_campagne_etal_n_etalonnage(self.type_de_saisie[1])
                
                
                
                #recuperation N de certificat :
                i = 0            
                while i < self.SpinBox_nbr_instruments.value() :                       
                    
                    #progress bar
                    self.progressBar.setValue(i)
                    
                    self.qlabel_campagne_etal_onglet_conclusion.setText(self.type_de_saisie[2])
                    
                    #preparation CE/CV
                    rapport_etalonnage = {} #dictionnaire ou on met les donnees afin de remplir le CE/CV
                    #Lists pour remplir le CE/CV
                    etalon = []
                    generateur = []
                    valeurs_etalon_corri = []
                    valeurs_instrum = []
                    valeurs_corrections = []
                    valeurs_U = []
                    valeurs_resolution = []
                    
                    
                    nom_fichier = str("resultat_" + "_instrum_"+ str(i+1)) #resultat__instrum_1
                    donnees = lecture_onglet_saisie(nom_fichier)
                    
                    identification_finale_instrument = str(donnees[0]["nom_instrument_"+str(i+1)])
                    id_final_instrument = db.recherche_id_instrument(identification_finale_instrument)
                    
                    #gestion mise à jour campagne
                        #list pour mettre à jour campagne:
                        #recuperation des id des instruments avant modification
                    list_id_instrum_avant_modif = db.recuperation_id_instrums_table_campagne_etalonnage_temp(id_campagne)
                    while len(list_id_instrum_avant_modif)!= 20:                    
                        list_id_instrum_avant_modif.append(0)
                    
                    list_finale_instruments = [ele if ele != id_ancien_instrum else id_final_instrument for ele in list_id_instrum_avant_modif]
 
                    db.mise_a_jour_table_campagne_etalonnage_temp(nom_campagne, date_etal, list_finale_instruments, id_campagne)
                    
                    #champs de la base :
                    nbr_pt_etalonnage = str("'"+str(self.SpinBox_nbr_pt_etal.value())+"'")           
                    identification_instrum = str("'"+str(donnees[0]["nom_instrument_"+str(i+1)])+"'")
                    rapport_etalonnage["identification_instrument"] = str(donnees[0]["nom_instrument_"+str(i+1)])
                    rapport_etalonnage["renseignement_complementaire"] = config_etal[i+1]["renseignement_complementaire"]
                    
                    rapport_etalonnage["Etat_reception"] = config_etal[1]["Etat_reception"]
                    etat_reception = (config_etal[1]["Etat_reception"].replace('"', '\"')).replace('\'','\'\'')
                    
                    
                    #requete pour avoir le n°serie , le constructeur ,designation , type , code du client:
                    caracteristique_instrum = db.recherche_caract_instrument(identification_instrum)
                    
                    n_serie =  caracteristique_instrum[0]
                    rapport_etalonnage["n_serie"] = caracteristique_instrum[1]
                    rapport_etalonnage["constructeur"] = caracteristique_instrum[2]
                    rapport_etalonnage["designation"] = caracteristique_instrum[3]
                    rapport_etalonnage["type"] = caracteristique_instrum[4]
                    rapport_etalonnage["affectation"] = caracteristique_instrum[5]
                    code_client = caracteristique_instrum[6]
                   
                    #technicien
                    id_technicien = str(donnees[0]['operateur']+1)
                    
                    caract_technicien = db.recherche_technicien_depuis_id(id_technicien)
                   
                    technicien =  caract_technicien[0]
                    rapport_etalonnage["operateur"] = caract_technicien[1]
                    
                    #Recherche du client             
                    caract_client = db.recherche_client(code_client)
                    
                    rapport_etalonnage["societe"] = caract_client[0]
                    rapport_etalonnage["adresse"] = caract_client[1]
                    rapport_etalonnage["code_postal"] = caract_client[2]
                    rapport_etalonnage["ville"] = caract_client[3] 
                    
                    
                    #recherche n°inventaire (SAP)
                    n_sap = db.recherche_n_inventaire(identification_instrum, n_serie)
                    rapport_etalonnage["n_sap"] = n_sap
                    rapport_etalonnage["date_etalonnage"] = str(date.day())+"/"+str(date.month())+"/"+str(date.year())
                    
                     
                    compteur = 0
                    list_pt_etal = []
                    while compteur < self.SpinBox_nbr_pt_etal.value():
                        list_pt_etal.append(donnees[compteur]["temp_consig"])
                        
                        if donnees[compteur]["generateur"] == 3 or donnees[compteur]["generateur"]== 4:
                            nom_procedure = str("'"+"PDL/PIL/SUR/MET/MO/030"+"'")
                            rapport_etalonnage["milieu"] = "AIR" 
                            rapport_etalonnage["n_mode_operatoire"] = "030"
                        else:
                            nom_procedure = str("'"+"PDL/PIL/SUR/MET/MO/025"+"'")
                            rapport_etalonnage["milieu"] = "LIQUIDE" 
                            rapport_etalonnage["n_mode_operatoire"] = "025"
                            
                        compteur +=1
                    
                    rapport_etalonnage["temps_consignes"] =list_pt_etal
                    
                        #on va chercher le N° du certificat avec la fonction numero_certificat
                    num_doc = self.type_de_saisie[1]
                    annule_doc = db.recherche_doc_annule_remplace(self.type_de_saisie[1])
                    
                    if annule_doc == "Neant":
                        rapport_etalonnage["n_certificat"] = num_doc
                        nom_fichier_ce =  num_doc
                        num_cv =  num_doc + "V" 
                        nom_fichier_cv = num_doc + "V" 
                    else:
                        rapport_etalonnage["n_certificat"] = num_doc + "\n Annule et Remplace le document "+str(annule_doc)
                        nom_fichier_ce =  num_doc + " Annule et Remplace le document "+str(annule_doc)
                        num_cv =  num_doc + "V" + "\n Annule et Remplace le document "+str(annule_doc)+"V"
                        nom_fichier_cv = num_doc + "V" + " Annule et Remplace le document "+str(annule_doc)+"V"
                    
                    
                    phrase_qlabel = "Le certificat n° {} de l'instrument {} est en cours de mise à jour".format(rapport_etalonnage["n_certificat"], rapport_etalonnage["identification_instrument"] )
                    self.qlabel_information_onglet_conclusion.setText(phrase_qlabel)
                    
                   
                        #Test à ecrire savoir si cofrac ou non
                    bouton_radio = [self.radioButton_cofrac_1, self.radioButton_cofrac_2, self.radioButton_cofrac_3, self.radioButton_cofrac_4, 
                                            self.radioButton_cofrac_5, self.radioButton_cofrac_6, self.radioButton_cofrac_7, self.radioButton_cofrac_8, 
                                            self.radioButton_cofrac_9, self.radioButton_cofrac_10, self.radioButton_cofrac_11, self.radioButton_cofrac_12, 
                                            self.radioButton_cofrac_13, self.radioButton_cofrac_14, self.radioButton_cofrac_15, self.radioButton_cofrac_16, 
                                            self.radioButton_cofrac_17, self.radioButton_cofrac_18, self.radioButton_cofrac_19, self.radioButton_cofrac_20]
                    
                    if bouton_radio[i].isChecked() :
                        edit_ce = str("'"+"COFRAC"+"'")
                        edit_cv =str("'"+"COFRAC"+"'")
                    else :
                        edit_ce = str("'"+"NON_COFRAC"+"'")
                        edit_cv =str("'"+"NON_COFRAC"+"'")
        
                    
                    ville_etal = str("'"+"Le Mans"+"'")
                    domaine_accreditation = str("'"+"Temperature"+"'")
                    if edit_ce == str("'"+"COFRAC"+"'"):
                        numero_accreditation = str("'"+"2-1950"+"'")
                    else : 
                        numero_accreditation = str("'"+"Neant"+"'")
                    nbr_pages_annexe = str("'"+str(0)+"'")
                    laboratoire = str("'"+"Laboratoire de Metrologie"+"'")
                    
                    
#                    Mise à jour table Etalonnage Temp Administration :
                    #Requete pour mise à jour des donnees dans la table etalonnage_temp_administration
                    id_campagne_etal  = db.recherche_id_campagne_etal_n_etalonnage(num_doc)
                    
                    num_doc_enlist = [num_doc]

                    
                    #Recuperation de l'id_etalonnage_temp_administration :
                    id_etalonnage_temp_administration = db.recherche_id_etalonnage_admin_id_campagne_identification_instrum(id_campagne_etal, identification_ancien_instrum)
                    rapport_etalonnage["id_etal"] = id_etalonnage_temp_administration
                    
                    db.mise_a_jour_table_etalonnage_temp_administration(id_campagne_etal, identification_ancien_instrum, nbr_pt_etalonnage, 
                                                                                        n_sap, identification_instrum, date_etal, 
                                                                                        nom_procedure, num_doc_enlist, edit_ce, edit_cv, 
                                                                                        technicien, ville_etal, domaine_accreditation, 
                                                                                        numero_accreditation, nbr_pages_annexe, 
                                                                                        laboratoire, etat_reception)
                ###########################################################
                    #Mise à jour Table ETALONNAGE_MESURE
                            #champs : CORRECTION, NUM_DOCUMENT, IDENTIFICATION, NBR_PT_ETALONNAGE, NUM_PT_ETAL, NBR_MESURE_PAR_PT, NUM_MESURE, MESURE_ETAL_NC, MESURE_ETAL_C, MESURE_INSTRUM
                
                    
                
                    #gestion profondeur immersion            
                    immersion = []
                    qlabel_prof_immersion = [self.textEdit_prof_imm_instrum_1, self.textEdit_prof_imm_instrum_2, self.textEdit_prof_imm_instrum_3, 
                                                    self.textEdit_prof_imm_instrum_4, self.textEdit_prof_imm_instrum_5, self.textEdit_prof_imm_instrum_6, 
                                                    self.textEdit_prof_imm_instrum_7, self.textEdit_prof_imm_instrum_8, self.textEdit_prof_imm_instrum_9, 
                                                    self.textEdit_prof_imm_instrum_10, self.textEdit_prof_imm_instrum_11, self.textEdit_prof_imm_instrum_12, 
                                                    self.textEdit_prof_imm_instrum_13, self.textEdit_prof_imm_instrum_14, self.textEdit_prof_imm_instrum_15, 
                                                    self.textEdit_prof_imm_instrum_16, self.textEdit_prof_imm_instrum_17, self.textEdit_prof_imm_instrum_18, 
                                                    self.textEdit_prof_imm_instrum_19, self.textEdit_prof_imm_instrum_20 ]
                
                    for ele in qlabel_prof_immersion[i].toPlainText().split() : 
                        immersion.append(ele)            
                
    
                    #gestion resolutions instrument saisie
                    resolution_instrument = []
                    qlabel_resolution_instrum = [self.textEdit_resolution_instrum_1, self.textEdit_resolution_instrum_2, self.textEdit_resolution_instrum_3, 
                                                        self.textEdit_resolution_instrum_4, self.textEdit_resolution_instrum_5, self.textEdit_resolution_instrum_6, 
                                                        self.textEdit_resolution_instrum_7, self.textEdit_resolution_instrum_8, self.textEdit_resolution_instrum_9, 
                                                        self.textEdit_resolution_instrum_10, self.textEdit_resolution_instrum_11, self.textEdit_resolution_instrum_12, 
                                                        self.textEdit_resolution_instrum_13, self.textEdit_resolution_instrum_14, self.textEdit_resolution_instrum_15, 
                                                        self.textEdit_resolution_instrum_16, self.textEdit_resolution_instrum_17, self.textEdit_resolution_instrum_18, 
                                                        self.textEdit_resolution_instrum_19, self.textEdit_resolution_instrum_20 ]
                    
                    for ele in qlabel_resolution_instrum[i].toPlainText().split() : 
                        resolution_instrument.append(ele.replace(",", "."))  
                
                
                
                    j=0
                    while j < self.SpinBox_nbr_pt_etal.value() :
                        nbr_de_mesure = donnees[j]["nb_acquisition"]
                    
                        k = 0
                        while k < nbr_de_mesure :
                        
                            correction = "'"+str(donnees[j]["corrections_instrum_"+str(i+1)][k])+"'"
                            num_pt_etalonnage ="'"+str(donnees[j]["pt_etalonnag_n"])+"'"
                            nbr_mesure = "'"+str(donnees[j]["nb_acquisition"])+"'"
                            numero_mesure = "'"+str(k+1)+"'"
                            mesure_etal_non_corri = "'"+str(donnees[j]["mesures_etal_brute"][k])+"'"
                            mesure_etal_corri =  "'"+str(donnees[j]["mesures_etal_corri"][k])+"'"
                            mesure_instrum = "'"+str(donnees[j]["mesures_instrum_"+str(i+1)][k])+"'"
    
                            #Mise à jour des donnees de mesures dans la table ETALONNAGE_MESURE 
                        
                            #Recherche de l'id ETAL MESURES
                            id_etal_mesures = db.recherche_id_etalonnage_mesure_id_campagne_id_etal(id_campagne_etal, id_etalonnage_temp_administration, numero_mesure, num_pt_etalonnage)
                            
                            db.mise_a_jour_table_etalonnage_mesure(id_etalonnage_temp_administration, id_campagne_etal, 
                                                                                correction, num_doc_enlist, identification_instrum, nbr_pt_etalonnage,
                                                                                num_pt_etalonnage, nbr_mesure, numero_mesure, mesure_etal_non_corri,
                                                                                mesure_etal_corri, mesure_instrum, id_etal_mesures)
    
                                        
                            k += 1
                    ###############################################################
                    #Mise à jour Table ETALONNAGE_RESULTAT
                    
                                #Remplissage Table etalonnage resultats:
                    
                       #champs
                                    
                        nom_etalon = "'"+str(self.Combobox_etalon_select_2.itemText(donnees[j]["etalon"]))+"'"
                        etalon.append(str(self.Combobox_etalon_select_2.itemText(donnees[j]["etalon"])))
                    
                        nom_generateur = "'"+str(self.Combobox_generateur_select_2.itemText(donnees[j]["generateur"]))+"'"
                        generateur.append(str(self.Combobox_generateur_select_2.itemText(donnees[j]["generateur"])))
                    
                        autoechauffement = "'"+str(donnees[j]["auto_echauffement"])+"'"
                    
                        fuite_thermique ="'"+str(donnees[j]["fuite_thermique"])+"'" 
                    
                        resolution = "'"+str(donnees[j]["Resolution"])+"'" # Attention il sagit de l'incertitude et non de la reso de l 'appareil
                        valeurs_resolution.append(donnees[j]["Resolution"])
                    
                        stabilite = "'"+str(donnees[j]["STAB"])+"'"
                    
                        eiv = "'"+str(donnees[j]["EIV"])+"'"
                    
                        moyenne_etalon_non_corri = "'"+str(donnees[j]["moyenne_etal_brute"])+"'"
                                    
                        moyenne_etalon_corri =  "'"+str(donnees[j]["moyenne_etal_corri"])+"'"
                        valeurs_etalon_corri.append(donnees[j]["moyenne_etal_corri"])
                    
                        moyenne_instrum = "'"+str(donnees[j]["moyenne_instrum_"+str(i+1)])+"'"
                        valeurs_instrum.append(donnees[j]["moyenne_instrum_"+str(i+1)])
                    
                        moyenne_correction = "'"+str(donnees[j]["moyenne_corrections_instrum_"+str(i+1)])+"'"
                        valeurs_corrections.append(donnees[j]["moyenne_corrections_instrum_"+str(i+1)])
                    
                        U = "'"+str(donnees[j]["U"])+"'"
                        valeurs_U.append(donnees[j]["U"])
                    
                        ecart_type = "'"+str(donnees[j]["ecart_type"])+"'"
                    
                        temp_consigne = donnees[j]["temp_consig"]
                    
                        valeur_immersion = immersion[j]
                    
                        valeur_resolution_instrument = resolution_instrument[j]
    
                        #insertion dans la table:
                    
                        id_etal_result = db.recherche_id_etal_result_table_etalonnage_resultat(id_campagne_etal, id_etalonnage_temp_administration)
                        db.mis_a_jour_table_etalonnage_resultat(temp_consigne, id_campagne_etal, id_etalonnage_temp_administration,
                                                                        nom_etalon, nom_generateur, autoechauffement , fuite_thermique, 
                                                                        resolution, stabilite, eiv, num_doc_enlist, identification_instrum, date_etal, 
                                                                        nbr_pt_etalonnage, num_pt_etalonnage, moyenne_etalon_non_corri, 
                                                                        moyenne_etalon_corri, moyenne_instrum, moyenne_correction, U, 
                                                                        ecart_type, valeur_immersion, valeur_resolution_instrument, id_etal_result[j])
    
                        j += 1
                
    #            resolution = lecture_onglet_configuration(self)            
                    rapport_etalonnage["resolution"] =  resolution_instrument[0] # resolution instrument
                
                    rapport_etalonnage["nbr_pt_etalonnage"] = self.SpinBox_nbr_pt_etal.value()
                    rapport_etalonnage["etalon"] = list(set(etalon)) # set : on crée un set qui permet d'eliminer les doublons
                    rapport_etalonnage["generateur"] = list(set(generateur)) # set : on crée un set qui permet d'eliminer les doublons
                    rapport_etalonnage["moyenne_etalon_corri"] = valeurs_etalon_corri
                    rapport_etalonnage["moyenne_instrum"] = valeurs_instrum
                    rapport_etalonnage["moyenne_correction"] =valeurs_corrections
                    rapport_etalonnage["U"] = valeurs_U
                    rapport_etalonnage["immersion"] = immersion
                
                    #Gestion du polynome correction =f(Tlue) : et enregistrement dans BDD            
                    if self.SpinBox_nbr_pt_etal.value() < 5:
                        ordre_poly = 1
                        poly = calcul_polynome (valeurs_instrum,valeurs_corrections, ordre_poly)
                        a = poly[0]
                        b = poly[1]
                        c = 0
                    else:
                        ordre_poly = 2
                        poly = calcul_polynome (valeurs_instrum,valeurs_corrections, ordre_poly)
                        a = poly[0]
                        b = poly[1]
                        c = poly[2]
                
                    date_du_jour = datetime.datetime.today()
                    date_du_jour_mise_en_forme = "'{}-{}-{}'".format(date_du_jour.year, date_du_jour.month, date_du_jour.day)
                    archivage = "'N'"
                
                    #mise à jour la table polynome
                    
                    if caracteristique_instrum[3] == "Etalon":
                        reponse = QMessageBox.question(self, 
                                    self.trUtf8("Information"), 
                                    self.trUtf8("Voulez-vous mettre à jour le polynome de correction"), 
                                    QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
                                    
                        
                        
                        if reponse == QtGui.QMessageBox.Yes:
                            db.mise_a_jour_table_polynome(identification_instrum,date_du_jour_mise_en_forme, num_doc_enlist,
                                                        date_etal, ordre_poly,a, b, c, archivage)
                        else:
                            pass
                        
                                     
                    else:
                        db.mise_a_jour_table_polynome(identification_instrum,date_du_jour_mise_en_forme, num_doc_enlist,
                                                        date_etal, ordre_poly,a, b, c, archivage)
                
                    
                    list_textedit_conformite =  [self.textEdit_resultat_instru_1, self.textEdit_resultat_instru_2, self.textEdit_resultat_instru_3,
                                                        self.textEdit_resultat_instru_4, self.textEdit_resultat_instru_5, self.textEdit_resultat_instru_6, 
                                                        self.textEdit_resultat_instru_7, self.textEdit_resultat_instru_8, self.textEdit_resultat_instru_9, 
                                                        self.textEdit_resultat_instru_10, self.textEdit_resultat_instru_11, self.textEdit_resultat_instru_12, 
                                                        self.textEdit_resultat_instru_13, self.textEdit_resultat_instru_14, self.textEdit_resultat_instru_15,
                                                        self.textEdit_resultat_instru_16, self.textEdit_resultat_instru_17, self.textEdit_resultat_instru_18, 
                                                        self.textEdit_resultat_instru_19, self.textEdit_resultat_instru_20]
                                                    
                                                
                    list_combobox_emt = [self.comboBox_ref_emt_1, self.comboBox_ref_emt_2, self.comboBox_ref_emt_3, 
                                            self.comboBox_ref_emt_4, self.comboBox_ref_emt_5, self.comboBox_ref_emt_6, 
                                            self.comboBox_ref_emt_7, self.comboBox_ref_emt_8, self.comboBox_ref_emt_7, 
                                            self.comboBox_ref_emt_8, self.comboBox_ref_emt_9, self.comboBox_ref_emt_10, 
                                            self.comboBox_ref_emt_11, self.comboBox_ref_emt_12, self.comboBox_ref_emt_13, 
                                            self.comboBox_ref_emt_14, self.comboBox_ref_emt_15, self.comboBox_ref_emt_16, 
                                            self.comboBox_ref_emt_17, self.comboBox_ref_emt_18, self.comboBox_ref_emt_19, 
                                            self.comboBox_ref_emt_20]
                    #gestion textedit conformite:
                    resultat_conformite = []
                    for ele in list_textedit_conformite[i].toPlainText().split():                
                        resultat_conformite.append(ele)   
                    
                    if not resultat_conformite:
                     rapport_etalonnage["conformite"] = "Néant"
                    
                    else:   
                        if resultat_conformite[0] == "Non":
                            rapport_etalonnage["conformite"] = resultat_conformite[0] + " " + resultat_conformite[1]
                        else:
                            rapport_etalonnage["conformite"] = resultat_conformite[0]
                        
                
                    #Recuperation nom referentiel            
                    id_referentiel_conformite = list_combobox_emt[i].currentIndex() + 1
                    rapport_etalonnage["referentiel_emt"] = db.recherche_nom_ref_emt(id_referentiel_conformite)
               
                    #gestion des CE/CV et la sauvegarde 
                    certificat = RapportEtalonnage(edit_ce) #appel à la classe RapportEtalonnage
                    certificat.mise_en_forme_ce(rapport_etalonnage, dossier, nom_fichier_ce)
    
                    #CV :
                    edition_cv = gestion_edition_cv(self) #retourne une list avec true ou false en fct s'il faut ou non edité le cv
                    #on regarde ds la bdd s'il existe une declaration de conf si oui on a un tupple avec id_conformite, id_ref, conclusion
                    conformite_instrum = db.recherche_id_conformite_temp_resultat_id_etalonnage(id_etalonnage_temp_administration)
                   
    #               print(conformite_instrum)
                    
                    if edition_cv[i] == True:
                        if conformite_instrum[0] == None:
                            db.insertion_table_conformite_temp_resultat(rapport_etalonnage, date_etal)
                        else:
                            db.mise_a_jour_table_conformite_temp_resultat(conformite_instrum [0], rapport_etalonnage, date_etal)
                        constat = RapportEtalonnage(edit_ce)
                        nom_constat ="{}V".format(rapport_etalonnage["n_certificat"])
                        
                        if annule_doc == "Neant":
                            constat.mise_en_forme_cv(rapport_etalonnage, dossier, nom_constat)
                        else:
                            constat.mise_en_forme_cv_annule_remplace(rapport_etalonnage, dossier, num_cv, nom_fichier_cv)
                    i += 1
          
            #effacement des fichiers temporaire (present dans le dossier AppData) 
            
            path =os.path.abspath("AppData/")
          
            for ele in os.listdir(path):
                path_total = str(path + "/"+str(ele))
                
                if os.path.isfile(path_total): # verification qu'il s'agit bien de fichier
                    os.remove(path_total)
                
            self.close()
        else:
            pass
    @pyqtSlot()
    def on_actionAnnule_et_Remplace_triggered(self):
        """
        Slot documentation goes here.
        """
        
        #Selection du dossier pour la sauvegarde des certificats:
        
        dossier = QFileDialog.getExistingDirectory(None ,  "Selectionner le dossier de sauvegarde des Rapports", 'y:/1.METROLOGIE/0.ARCHIVES ETALONNAGE-VERIFICATIONS/1-TEMPERATURE/')
        if dossier != "":
            #gestion onglete conclusion
            self.tabWidget.setCurrentIndex(4)        
            self.progressBar.setMaximum(self.SpinBox_nbr_instruments.value() )
            self.progressBar.setMinimum(0)
            self.progressBar.setValue(0)
            
            phrase_qlabel = "Connexion à la base pour création n°campagne"
            self.qlabel_information_onglet_conclusion.setText(phrase_qlabel)
            #Acces bdd
            db = GestionBdd('db')
            db.reconnexion()
            
            ##################################################
            #constantes reutilisables
            date = self.calendarWidget.selectedDate()
            date_etal = date.toString('yyyy-MM-dd')
            date_etal_fr = date.toString('dd/MM/yyyy')
            ###################################################
            #Remplissage Table CAMPAGNE_ETALONNAGE_TEMP
            phrase_qlabel = "Creation n°campagne"
            self.qlabel_information_onglet_conclusion.setText(phrase_qlabel)
            
            nom_fichier = "configuration"
            config_etal = lecture_onglet_saisie(nom_fichier)
            
            caracteristiques_campagne = numero_campagne_etal(self) #attention numero_campagne renvoie le nom et l'id insere dans la bdd
    #        print(caracteristiques_campagne)
            nom_campagne = caracteristiques_campagne[0]        
            id_campagne_etal = caracteristiques_campagne[1]
    
                #affichage pour l'utilisateur nom campagne
            
            QMessageBox.information(self, 
                        self.trUtf8("Campagne d'etalonnage"), 
                        self.trUtf8("Merci de relever la reference de la campagne d'etalonnage : {}".format(nom_campagne)))
            
            #Gestion onglet conclusion :
            information_campagne = "Campagne d'étalonnage n° {}".format(nom_campagne)        
            self.qlabel_campagne_etal_onglet_conclusion.setText(information_campagne)
            
            
            
            
            list_instrum = []
            i = 0
            while i < self.SpinBox_nbr_instruments.value():
                list_instrum.append(str(config_etal[i+1]["nom_instrum"]))
                
                i+=1
                
            # on recherche les ID_instrument
            
            id_instrum = []
            for ele in list_instrum:
                id = db.recherche_id_instrument(ele)
                id_instrum.append(id)
    
            
            compteur = 20 - self.SpinBox_nbr_instruments.value()
                   
            i= 0
            while i < compteur: 
                id_instrum.append(0)
                i+=1
            
            #Insertion donnees table campagne etal:
                   
            phrase_qlabel = "sauvegarde campagne etalonnage"
            self.qlabel_information_onglet_conclusion.setText(phrase_qlabel)
            
            db.mise_a_jour_table_campagne_etalonnage_temp(nom_campagne, date_etal, id_instrum, id_campagne_etal)
    
    #        id_campagne_etal = db.insertion_table_campagne_etal_2(nom_campagne, date_etal, id_instrum)
    #        db.insertion_table_campagne_etal(nom_campagne, date_etal, id_instrum)
    #        
    #        # recuperation de l'id campagne etalonnage :        
    #        id_campagne_etal = db.recherche_id_campagne_etal()
                 
           
            #Remplissage table :ETALONNAGE_TEMP_ADMINISTRATION
            ############################################################################
                    
    
            
            #progress bar
            self.progressBar.setValue(1)
            
            
            #preparation CE/CV
            rapport_etalonnage = {} #dictionnaire ou on met les donnees afin de remplir le CE/CV
            #Lists pour remplir le CE/CV
            etalon = []
            generateur = []
            valeurs_etalon_corri = []
            valeurs_instrum = []
            valeurs_corrections = []
            valeurs_U = []
            valeurs_resolution = []
            
            
            nom_fichier = str("resultat_" + "_instrum_"+ str(1)) #resultat__instrum_1
            donnees = lecture_onglet_saisie(nom_fichier)
            
            
            #champs de la base :
            nbr_pt_etalonnage = str("'"+str(self.SpinBox_nbr_pt_etal.value())+"'")           
            identification_instrum = str("'"+str(donnees[0]["nom_instrument_"+str(1)])+"'")
            rapport_etalonnage["identification_instrument"] = str(donnees[0]["nom_instrument_"+str(1)])
            rapport_etalonnage["renseignement_complementaire"] = config_etal[1]["renseignement_complementaire"]
            rapport_etalonnage["Etat_reception"] = config_etal[1]["Etat_reception"]
            #requete pour avoir le n°serie , le constructeur ,designation , type , code du client:
            caracteristique_instrum = db.recherche_caract_instrument(identification_instrum)
            
            etat_reception = (config_etal[1]["Etat_reception"].replace('"', '\"')).replace('\'','\'\'')
            n_serie =  caracteristique_instrum[0]
            n_seris_bis = caracteristique_instrum[1]
            rapport_etalonnage["n_serie"] = caracteristique_instrum[1]
            rapport_etalonnage["constructeur"] = caracteristique_instrum[2]
            rapport_etalonnage["designation"] = caracteristique_instrum[3]
            rapport_etalonnage["type"] = caracteristique_instrum[4]
            rapport_etalonnage["affectation"] = caracteristique_instrum[5]
    #            rapport_etalonnage["renseignement_complementaire"] = donnnee
            code_client = caracteristique_instrum[6]
           
            #technicien
            id_technicien = str(donnees[0]['operateur']+1)
            
            caract_technicien = db.recherche_technicien_depuis_id(id_technicien)
           
            technicien =  caract_technicien[0]
            rapport_etalonnage["operateur"] = caract_technicien[1]
            
            #Recherche du client             
            caract_client = db.recherche_client(code_client)
            
            rapport_etalonnage["societe"] = caract_client[0]
            rapport_etalonnage["adresse"] = caract_client[1]
            rapport_etalonnage["code_postal"] = caract_client[2]
            rapport_etalonnage["ville"] = caract_client[3] 
            
            
            #recherche n°inventaire (SAP)
            n_sap = db.recherche_n_inventaire(identification_instrum, n_seris_bis)
            
            if isinstance(n_sap, str):
                rapport_etalonnage["n_sap"] = n_sap 
            else:
                n_sap = "Neant"
                rapport_etalonnage["n_sap"] = "Neant"          
            
            
            rapport_etalonnage["date_etalonnage"] = date_etal_fr #str(date.day())+"/"+str(date.month())+"/"+str(date.year())
            
             
            compteur = 0
            list_pt_etal = []
            while compteur < self.SpinBox_nbr_pt_etal.value():
                list_pt_etal.append(donnees[compteur]["temp_consig"])
                
                if donnees[compteur]["generateur"] == 3 or donnees[compteur]["generateur"]== 4:
                    nom_procedure = str("'"+"PDL/PIL/SUR/MET/MO/030"+"'")
                    rapport_etalonnage["milieu"] = "AIR" 
                    rapport_etalonnage["n_mode_operatoire"] = "030"
                else:
                    nom_procedure = str("'"+"PDL/PIL/SUR/MET/MO/025"+"'")
                    rapport_etalonnage["milieu"] = "LIQUIDE" 
                    rapport_etalonnage["n_mode_operatoire"] = "025"
                    
                compteur +=1
            
            rapport_etalonnage["temps_consignes"] =list_pt_etal
        
            #on va chercher le N° du certificat avec la fonction numero_certificat
                #on va chercher le N° du certificat avec la fonction numero_certificat
            caracteristique_etalonnage = numero_certificat(self) #numero certf renvoie le nom et l'id
            num_doc = str(caracteristique_etalonnage[0])
            num_doc_annul = str(self.type_de_saisie[1])
    
            rapport_etalonnage["id_etal"] = caracteristique_etalonnage[1]
            id_etalonnage_temp_administration = caracteristique_etalonnage[1]
        
            rapport_etalonnage["n_certificat"] = num_doc + "\n Annule et Remplace le document "+str(self.type_de_saisie[1])
                
            phrase_qlabel = "Le certificat n° {} de l'instrument {} est en cours de création".format(rapport_etalonnage["n_certificat"], rapport_etalonnage["identification_instrument"] )
            self.qlabel_information_onglet_conclusion.setText(phrase_qlabel)
           
            nom_fichier_ce =  num_doc + " Annule et Remplace le document "+str(self.type_de_saisie[1])
            num_cv =  num_doc + "V" + "\n Annule et Remplace le document "+str(self.type_de_saisie[1])+"V"
            nom_fichier_cv = num_doc + "V" + " Annule et Remplace le document "+str(self.type_de_saisie[1])+"V"
           
                #Test à ecrire savoir si cofrac ou non
            bouton_radio = [self.radioButton_cofrac_1, self.radioButton_cofrac_2, self.radioButton_cofrac_3, self.radioButton_cofrac_4, 
                                    self.radioButton_cofrac_5, self.radioButton_cofrac_6, self.radioButton_cofrac_7, self.radioButton_cofrac_8, 
                                    self.radioButton_cofrac_9, self.radioButton_cofrac_10, self.radioButton_cofrac_11, self.radioButton_cofrac_12, 
                                    self.radioButton_cofrac_13, self.radioButton_cofrac_14, self.radioButton_cofrac_15, self.radioButton_cofrac_16, 
                                    self.radioButton_cofrac_17, self.radioButton_cofrac_18, self.radioButton_cofrac_19, self.radioButton_cofrac_20]
            
            if bouton_radio[0].isChecked() :
                edit_ce = str("'"+"COFRAC"+"'")
                edit_cv =str("'"+"COFRAC"+"'")
            else :
                edit_ce = str("'"+"NON_COFRAC"+"'")
                edit_cv =str("'"+"NON_COFRAC"+"'")
    
            
            ville_etal = str("'"+"Le Mans"+"'")
            domaine_accreditation = str("'"+"Temperature"+"'")
            if edit_ce == str("'"+"COFRAC"+"'"):
                numero_accreditation = str("'"+"2-1950"+"'")
            else : 
                numero_accreditation = str("'"+"Neant"+"'")
            nbr_pages_annexe = str("'"+str(0)+"'")
            laboratoire = str("'"+"Laboratoire de Metrologie"+"'")
            
            #Requete pour envoyer les donnees dans la table etalonnage_temp_administration
            
             #Requete pour envoyer les donnees dans la table etalonnage_temp_administration
            
            db.mise_a_jour_table_etalonnage_temp_administration_annule_remplace_2(id_campagne_etal, nbr_pt_etalonnage, 
                                                                                    n_sap, identification_instrum, date_etal, 
                                                                                    nom_procedure, num_doc, edit_ce, edit_cv, 
                                                                                    technicien, ville_etal, domaine_accreditation, 
                                                                                    numero_accreditation, nbr_pages_annexe, 
                                                                                    laboratoire, num_doc_annul, etat_reception, id_etalonnage_temp_administration)
        
        
        ##################################################################################
        #Remplissage table etalonnage mesure
        
            #champs : CORRECTION, NUM_DOCUMENT, IDENTIFICATION, NBR_PT_ETALONNAGE, NUM_PT_ETAL, NBR_MESURE_PAR_PT, NUM_MESURE, MESURE_ETAL_NC, MESURE_ETAL_C, MESURE_INSTRUM
            
            #gestion profondeur immersion            
            immersion = []
            qlabel_prof_immersion = [self.textEdit_prof_imm_instrum_1, self.textEdit_prof_imm_instrum_2, self.textEdit_prof_imm_instrum_3, 
                                                self.textEdit_prof_imm_instrum_4, self.textEdit_prof_imm_instrum_5, self.textEdit_prof_imm_instrum_6, 
                                                self.textEdit_prof_imm_instrum_7, self.textEdit_prof_imm_instrum_8, self.textEdit_prof_imm_instrum_9, 
                                                self.textEdit_prof_imm_instrum_10, self.textEdit_prof_imm_instrum_11, self.textEdit_prof_imm_instrum_12, 
                                                self.textEdit_prof_imm_instrum_13, self.textEdit_prof_imm_instrum_14, self.textEdit_prof_imm_instrum_15, 
                                                self.textEdit_prof_imm_instrum_16, self.textEdit_prof_imm_instrum_17, self.textEdit_prof_imm_instrum_18, 
                                                self.textEdit_prof_imm_instrum_19, self.textEdit_prof_imm_instrum_20 ]
            
            for ele in qlabel_prof_immersion[0].toPlainText().split() : 
                immersion.append(ele)            
            
    
            #gestion resolutions instrument saisie
            resolution_instrument = []
            qlabel_resolution_instrum = [self.textEdit_resolution_instrum_1, self.textEdit_resolution_instrum_2, self.textEdit_resolution_instrum_3, 
                                                    self.textEdit_resolution_instrum_4, self.textEdit_resolution_instrum_5, self.textEdit_resolution_instrum_6, 
                                                    self.textEdit_resolution_instrum_7, self.textEdit_resolution_instrum_8, self.textEdit_resolution_instrum_9, 
                                                    self.textEdit_resolution_instrum_10, self.textEdit_resolution_instrum_11, self.textEdit_resolution_instrum_12, 
                                                    self.textEdit_resolution_instrum_13, self.textEdit_resolution_instrum_14, self.textEdit_resolution_instrum_15, 
                                                    self.textEdit_resolution_instrum_16, self.textEdit_resolution_instrum_17, self.textEdit_resolution_instrum_18, 
                                                    self.textEdit_resolution_instrum_19, self.textEdit_resolution_instrum_20 ]
                
            for ele in qlabel_resolution_instrum[0].toPlainText().split() : 
                resolution_instrument.append(ele.replace(",", "."))  
            
                    
            j=0
            while j < self.SpinBox_nbr_pt_etal.value() :
                nbr_de_mesure = donnees[j]["nb_acquisition"]
                
                k = 0
                while k < nbr_de_mesure :
                    
                    correction = "'"+str(donnees[j]["corrections_instrum_"+str(1)][k])+"'"
                    num_pt_etalonnage ="'"+str(donnees[j]["pt_etalonnag_n"])+"'"
                    nbr_mesure = "'"+str(donnees[j]["nb_acquisition"])+"'"
                    numero_mesure = "'"+str(k+1)+"'"
                    mesure_etal_non_corri = "'"+str(donnees[j]["mesures_etal_brute"][k])+"'"
                    mesure_etal_corri =  "'"+str(donnees[j]["mesures_etal_corri"][k])+"'"
                    mesure_instrum = "'"+str(donnees[j]["mesures_instrum_"+str(1)][k])+"'"
    
                    #Insertion des donnees de mesures dans la table ETALONNAGE_MESURE 
                    
                    db. insertion_table_etalonnage_mesure(id_etalonnage_temp_administration, id_campagne_etal, 
                                                                            correction, num_doc, identification_instrum, nbr_pt_etalonnage,
                                                                            num_pt_etalonnage, nbr_mesure, numero_mesure, mesure_etal_non_corri,
                                                                            mesure_etal_corri, mesure_instrum)
    
                                    
                    k += 1
    #######################################################################################################################
            #Remplissage Table etalonnage resultats:
                
                #champs
                                
                nom_etalon = "'"+str(self.Combobox_etalon_select_2.itemText(donnees[j]["etalon"]))+"'"
                etalon.append(str(self.Combobox_etalon_select_2.itemText(donnees[j]["etalon"])))
                
                nom_generateur = "'"+str(self.Combobox_generateur_select_2.itemText(donnees[j]["generateur"]))+"'"
                generateur.append(str(self.Combobox_generateur_select_2.itemText(donnees[j]["generateur"])))
                
                autoechauffement = "'"+str(donnees[j]["auto_echauffement"])+"'"
                
                fuite_thermique ="'"+str(donnees[j]["fuite_thermique"])+"'" 
                
                resolution = "'"+str(donnees[j]["Resolution"])+"'" # Attention il sagit de l'incertitude et non de la reso de l 'appareil
                valeurs_resolution.append(donnees[j]["Resolution"])
                
                stabilite = "'"+str(donnees[j]["STAB"])+"'"
                
                eiv = "'"+str(donnees[j]["EIV"])+"'"
                
                moyenne_etalon_non_corri = "'"+str(donnees[j]["moyenne_etal_brute"])+"'"
                                
                moyenne_etalon_corri =  "'"+str(donnees[j]["moyenne_etal_corri"])+"'"
                valeurs_etalon_corri.append(donnees[j]["moyenne_etal_corri"])
                
                moyenne_instrum = "'"+str(donnees[j]["moyenne_instrum_"+str(1)])+"'"
                valeurs_instrum.append(donnees[j]["moyenne_instrum_"+str(1)])
                
                moyenne_correction = "'"+str(donnees[j]["moyenne_corrections_instrum_"+str(1)])+"'"
                valeurs_corrections.append(donnees[j]["moyenne_corrections_instrum_"+str(1)])
                
                U = "'"+str(donnees[j]["U"])+"'"
                valeurs_U.append(donnees[j]["U"])
                
                ecart_type = "'"+str(donnees[j]["ecart_type"])+"'"
                
                temp_consigne = donnees[j]["temp_consig"]
                
                valeur_immersion = immersion[j]
                
                valeur_resolution_instrument = resolution_instrument[j]
    
                #insertion dans la table:
                db. insertion_table_etalonnage_resultat(temp_consigne, id_campagne_etal, id_etalonnage_temp_administration,
                                                                    nom_etalon, nom_generateur, autoechauffement , fuite_thermique, 
                                                                    resolution, stabilite, eiv, num_doc, identification_instrum, date_etal, 
                                                                    nbr_pt_etalonnage, num_pt_etalonnage, moyenne_etalon_non_corri, 
                                                                    moyenne_etalon_corri, moyenne_instrum, moyenne_correction, U, 
                                                                    ecart_type, valeur_immersion, valeur_resolution_instrument)
    
                j += 1
            
    #            resolution =lecture_onglet_configuration(self)            
            rapport_etalonnage["resolution"] =  resolution_instrument[0] # resolution instrument
            
            rapport_etalonnage["nbr_pt_etalonnage"] = self.SpinBox_nbr_pt_etal.value()
            rapport_etalonnage["etalon"] = list(set(etalon)) # set : on crée un set qui permet d'eliminer les doublons
            rapport_etalonnage["generateur"] = list(set(generateur)) # set : on crée un set qui permet d'eliminer les doublons
            rapport_etalonnage["moyenne_etalon_corri"] = valeurs_etalon_corri
            rapport_etalonnage["moyenne_instrum"] = valeurs_instrum
            rapport_etalonnage["moyenne_correction"] =valeurs_corrections
            rapport_etalonnage["U"] = valeurs_U
            rapport_etalonnage["immersion"] = immersion
            
            #Gestion du polynome correction =f(Tlue) : et enregistrement dans BDD            
            if self.SpinBox_nbr_pt_etal.value() < 5:
                ordre_poly = 1
                poly = calcul_polynome (valeurs_instrum,valeurs_corrections, ordre_poly)
                a = poly[0]
                b = poly[1]
                c = 0
            else:
                ordre_poly = 2
                poly = calcul_polynome (valeurs_instrum,valeurs_corrections, ordre_poly)
                a = poly[0]
                b = poly[1]
                c = poly[2]
            
            date_du_jour = datetime.datetime.today()
            date_du_jour_mise_en_forme = "'{}-{}-{}'".format(date_du_jour.year, date_du_jour.month, date_du_jour.day)
            archivage = "'N'"
            
            #insertion dans la table polynome
            db.insertion_table_polynome(identification_instrum,date_du_jour_mise_en_forme, num_doc, \
                                                    date_etal, ordre_poly,a, b, c, archivage)
            
            
    
            
            list_textedit_conformite =  [self.textEdit_resultat_instru_1, self.textEdit_resultat_instru_2, self.textEdit_resultat_instru_3,
                                                    self.textEdit_resultat_instru_4, self.textEdit_resultat_instru_5, self.textEdit_resultat_instru_6, 
                                                    self.textEdit_resultat_instru_7, self.textEdit_resultat_instru_8, self.textEdit_resultat_instru_9, 
                                                    self.textEdit_resultat_instru_10, self.textEdit_resultat_instru_11, self.textEdit_resultat_instru_12, 
                                                    self.textEdit_resultat_instru_13, self.textEdit_resultat_instru_14, self.textEdit_resultat_instru_15, 
                                                    self.textEdit_resultat_instru_16, self.textEdit_resultat_instru_17, self.textEdit_resultat_instru_18, 
                                                    self.textEdit_resultat_instru_19, self.textEdit_resultat_instru_20]
                                                
                                            
            list_combobox_emt = [self.comboBox_ref_emt_1, self.comboBox_ref_emt_2, self.comboBox_ref_emt_3, 
                                        self.comboBox_ref_emt_4, self.comboBox_ref_emt_5, self.comboBox_ref_emt_6, 
                                        self.comboBox_ref_emt_7, self.comboBox_ref_emt_8, self.comboBox_ref_emt_7, 
                                        self.comboBox_ref_emt_8, self.comboBox_ref_emt_9, self.comboBox_ref_emt_10, 
                                        self.comboBox_ref_emt_11, self.comboBox_ref_emt_12, self.comboBox_ref_emt_13, 
                                        self.comboBox_ref_emt_14, self.comboBox_ref_emt_15, self.comboBox_ref_emt_16, 
                                        self.comboBox_ref_emt_17, self.comboBox_ref_emt_18, self.comboBox_ref_emt_19, 
                                        self.comboBox_ref_emt_20]
            #gestion textedit conformite:
            resultat_conformite = []
            for ele in list_textedit_conformite[0].toPlainText().split():                
                resultat_conformite.append(ele)
            
            if not resultat_conformite:
                rapport_etalonnage["conformite"] = "Néant"    
            else:
                if resultat_conformite[0] == "Non":
                    rapport_etalonnage["conformite"] = resultat_conformite[0] + " " + resultat_conformite[1]
                else:
                    rapport_etalonnage["conformite"] = resultat_conformite[0]
            
            
            self.progressBar.setValue(1)
            
           
            #Recuperation nom referentiel            
            nom_referentiel_conformite = list_combobox_emt[0].currentText()
            id_referentiel_conformite = db.recherche_id_ref_emt(nom_referentiel_conformite)
            rapport_etalonnage["referentiel_emt"] = db.recherche_nom_ref_emt(id_referentiel_conformite)
                 
            #gestion des CE/CV et la sauvegarde 
    #        print(rapport_etalonnage)
            certificat = RapportEtalonnage(edit_ce) #appel à la classe RapportEtalonnage
            certificat.mise_en_forme_ce(rapport_etalonnage, dossier, nom_fichier_ce)
    
            #CV :
            if self.checkBox.isChecked():        
                db.insertion_table_conformite_temp_resultat(rapport_etalonnage, date_etal)
                constat = RapportEtalonnage(edit_ce)
                nom_constat ="{}V".format(rapport_etalonnage["n_certificat"])
                constat.mise_en_forme_cv_annule_remplace(rapport_etalonnage, dossier, num_cv, nom_fichier_cv)
            self.progressBar.setValue(2)
            self.close()
        else:
            pass
        
    @pyqtSlot()
    def on_actionEnregistrer_triggered(self):
        """
        Permet la sauvegarde de l'ensemble de l'etalonnage dans la base et de generer les CE/CV"""
        
        #Selection du dossier pour la sauvegarde des certificats:
        
        dossier = QFileDialog.getExistingDirectory(None ,  "Selectionner le dossier de sauvegarde des Rapports", 'y:/1.METROLOGIE/0.ARCHIVES ETALONNAGE-VERIFICATIONS/1-TEMPERATURE/')
        if dossier !="":
            #gestion onglete conclusion
            self.tabWidget.setCurrentIndex(4)        
            self.progressBar.setMaximum(self.SpinBox_nbr_instruments.value() )
            self.progressBar.setMinimum(0)
            self.progressBar.setValue(0)
            
            phrase_qlabel = "Connexion à la base pour création n°campagne"
            self.qlabel_information_onglet_conclusion.setText(phrase_qlabel)
            #Acces bdd
            db = GestionBdd('db')
            db.reconnexion()
            
            ##################################################
            #constantes reutilisables
            date = self.calendarWidget.selectedDate()
            date_etal = date.toString('yyyy-MM-dd')
            date_etal_fr = date.toString('dd/MM/yyyy')
            ###################################################
            #Remplissage Table CAMPAGNE_ETALONNAGE_TEMP
            phrase_qlabel = "Creation n°campagne"
            self.qlabel_information_onglet_conclusion.setText(phrase_qlabel)
            
            nom_fichier = "configuration"
            config_etal = lecture_onglet_saisie(nom_fichier)
            
            caracteristiques_campagne = numero_campagne_etal(self) #attention numero_campagne renvoie le nom et l'id insere dans la bdd
    #        print(caracteristiques_campagne)
            nom_campagne = caracteristiques_campagne[0]        
            id_campagne_etal = caracteristiques_campagne[1]
    
                #affichage pour l'utilisateur nom campagne
            
            QMessageBox.information(self, 
                        self.trUtf8("Campagne d'etalonnage"), 
                        self.trUtf8("Merci de relever la reference de la campagne d'etalonnage : {}".format(nom_campagne)))
            
            #Gestion onglet conclusion :
            information_campagne = "Campagne d'étalonnage n° {}".format(nom_campagne)        
            self.qlabel_campagne_etal_onglet_conclusion.setText(information_campagne)
            
            
            
            
            list_instrum = []
            i = 0
            while i < self.SpinBox_nbr_instruments.value():
                list_instrum.append(str(config_etal[i+1]["nom_instrum"]))
                
                i+=1
                
            # on recherche les ID_instrument
            
            id_instrum = []
            for ele in list_instrum:
                id = db.recherche_id_instrument(ele)
                id_instrum.append(id)
    
            
            compteur = 20 - self.SpinBox_nbr_instruments.value()
                   
            i= 0
            while i < compteur: 
                id_instrum.append(0)
                i+=1
            
            #Insertion donnees table campagne etal:
                   
            phrase_qlabel = "sauvegarde campagne etalonnage"
            self.qlabel_information_onglet_conclusion.setText(phrase_qlabel)
            
            db.mise_a_jour_table_campagne_etalonnage_temp(nom_campagne, date_etal, id_instrum, id_campagne_etal)
    
    #        id_campagne_etal = db.insertion_table_campagne_etal_2(nom_campagne, date_etal, id_instrum)
    #        db.insertion_table_campagne_etal(nom_campagne, date_etal, id_instrum)
    #        
    #        # recuperation de l'id campagne etalonnage :        
    #        id_campagne_etal = db.recherche_id_campagne_etal()
                 
           
            #Remplissage table :ETALONNAGE_TEMP_ADMINISTRATION
            ############################################################################
                    
            i = 0
            while i < self.SpinBox_nbr_instruments.value() :
                
                #progress bar
                self.progressBar.setValue(i)
                
                
                #preparation CE/CV
                rapport_etalonnage = {} #dictionnaire ou on met les donnees afin de remplir le CE/CV
                #Lists pour remplir le CE/CV
                etalon = []
                generateur = []
                valeurs_etalon_corri = []
                valeurs_instrum = []
                valeurs_corrections = []
                valeurs_U = []
                valeurs_resolution = []
                
                
                nom_fichier = str("resultat_" + "_instrum_"+ str(i+1)) #resultat__instrum_1
                donnees = lecture_onglet_saisie(nom_fichier)
                
                
                #champs de la base :
                nbr_pt_etalonnage = str("'"+str(self.SpinBox_nbr_pt_etal.value())+"'")           
                identification_instrum = str("'"+str(donnees[0]["nom_instrument_"+str(i+1)])+"'")
                rapport_etalonnage["identification_instrument"] = str(donnees[0]["nom_instrument_"+str(i+1)])
                rapport_etalonnage["renseignement_complementaire"] = config_etal[i+1]["renseignement_complementaire"]
                rapport_etalonnage["Etat_reception"] = config_etal[i+1]["Etat_reception"]
                #requete pour avoir le n°serie , le constructeur ,designation , type , code du client:
                caracteristique_instrum = db.recherche_caract_instrument(identification_instrum)
                
                etat_reception = (config_etal[i+1]["Etat_reception"].replace('"', '\"')).replace('\'','\'\'')
                n_serie =  caracteristique_instrum[0]
                n_seris_bis = caracteristique_instrum[1]
                rapport_etalonnage["n_serie"] = caracteristique_instrum[1]
                rapport_etalonnage["constructeur"] = caracteristique_instrum[2]
                rapport_etalonnage["designation"] = caracteristique_instrum[3]
                rapport_etalonnage["type"] = caracteristique_instrum[4]
                rapport_etalonnage["affectation"] = caracteristique_instrum[5]
    #            rapport_etalonnage["renseignement_complementaire"] = donnnee
                code_client = caracteristique_instrum[6]
               
                #technicien
                id_technicien = str(donnees[0]['operateur']+1)
                
                caract_technicien = db.recherche_technicien_depuis_id(id_technicien)
               
                technicien =  caract_technicien[0]
                rapport_etalonnage["operateur"] = caract_technicien[1]
                
                #Recherche du client             
                caract_client = db.recherche_client(code_client)
                
                rapport_etalonnage["societe"] = caract_client[0]
                rapport_etalonnage["adresse"] = caract_client[1]
                rapport_etalonnage["code_postal"] = caract_client[2]
                rapport_etalonnage["ville"] = caract_client[3] 
                
                
                #recherche n°inventaire (SAP)
                n_sap = db.recherche_n_inventaire(identification_instrum, n_seris_bis)
                
                if isinstance(n_sap, str):
                    rapport_etalonnage["n_sap"] = n_sap 
                else:
                    n_sap = "Neant"
                    rapport_etalonnage["n_sap"] = "Neant"
                
                
                rapport_etalonnage["date_etalonnage"] = date_etal_fr #str(date.day())+"/"+str(date.month())+"/"+str(date.year())
                
                 
                compteur = 0
                list_pt_etal = []
                while compteur < self.SpinBox_nbr_pt_etal.value():
                    list_pt_etal.append(donnees[compteur]["temp_consig"])
                    
                    if donnees[compteur]["generateur"] == 3 or donnees[compteur]["generateur"]== 4:
                        nom_procedure = str("'"+"PDL/PIL/SUR/MET/MO/030"+"'")
                        rapport_etalonnage["milieu"] = "AIR" 
                        rapport_etalonnage["n_mode_operatoire"] = "030"
                    else:
                        nom_procedure = str("'"+"PDL/PIL/SUR/MET/MO/025"+"'")
                        rapport_etalonnage["milieu"] = "LIQUIDE" 
                        rapport_etalonnage["n_mode_operatoire"] = "025"
                        
                    compteur +=1
                
                rapport_etalonnage["temps_consignes"] =list_pt_etal
                
                    #on va chercher le N° du certificat avec la fonction numero_certificat
                caracteristique_etalonnage = numero_certificat(self) #numero certf renvoie le nom et l'id
                num_doc = str(caracteristique_etalonnage[0])
    
                rapport_etalonnage["id_etal"] = caracteristique_etalonnage[1]
                id_etalonnage_temp_administration = caracteristique_etalonnage[1]
    
                rapport_etalonnage["n_certificat"] = num_doc
                
                phrase_qlabel = "Le certificat n° {} de l'instrument {} est en cours de création".format(rapport_etalonnage["n_certificat"], rapport_etalonnage["identification_instrument"] )
                self.qlabel_information_onglet_conclusion.setText(phrase_qlabel)
                
               
                    #Test à ecrire savoir si cofrac ou non
                bouton_radio = [self.radioButton_cofrac_1, self.radioButton_cofrac_2, self.radioButton_cofrac_3, self.radioButton_cofrac_4, 
                                        self.radioButton_cofrac_5, self.radioButton_cofrac_6, self.radioButton_cofrac_7, self.radioButton_cofrac_8, 
                                        self.radioButton_cofrac_9, self.radioButton_cofrac_10, self.radioButton_cofrac_11, self.radioButton_cofrac_12, 
                                        self.radioButton_cofrac_13, self.radioButton_cofrac_14, self.radioButton_cofrac_15, self.radioButton_cofrac_16, 
                                        self.radioButton_cofrac_17, self.radioButton_cofrac_18, self.radioButton_cofrac_19, self.radioButton_cofrac_20]
                
                if bouton_radio[i].isChecked() :
                    edit_ce = str("'"+"COFRAC"+"'")
                    edit_cv =str("'"+"COFRAC"+"'")
                else :
                    edit_ce = str("'"+"NON_COFRAC"+"'")
                    edit_cv =str("'"+"NON_COFRAC"+"'")
    
                
                ville_etal = str("'"+"Le Mans"+"'")
                domaine_accreditation = str("'"+"Temperature"+"'")
                if edit_ce == str("'"+"COFRAC"+"'"):
                    numero_accreditation = str("'"+"2-1950"+"'")
                else : 
                    numero_accreditation = str("'"+"Neant"+"'")
                nbr_pages_annexe = str("'"+str(0)+"'")
                laboratoire = str("'"+"Laboratoire de Metrologie"+"'")
                
                #Requete pour envoyer les donnees dans la table etalonnage_temp_administration
                db.mise_a_jour_table_etalonnage_temp_administration_2(id_campagne_etal, nbr_pt_etalonnage, 
                                                                                        n_sap, identification_instrum, date_etal, 
                                                                                        nom_procedure, num_doc, edit_ce, edit_cv, 
                                                                                        technicien, ville_etal, domaine_accreditation, 
                                                                                        numero_accreditation, nbr_pages_annexe, 
                                                                                        laboratoire, etat_reception, id_etalonnage_temp_administration)
                
                
                
    #            id_etalonnage_temp_administration = db.insertion_table_etalonnage_temp_administration_2(id_campagne_etal, nbr_pt_etalonnage, 
    #                                                                                    n_sap, identification_instrum, date_etal, 
    #                                                                                    nom_procedure, num_doc, edit_ce, edit_cv, 
    #                                                                                    technicien, ville_etal, domaine_accreditation, 
    #                                                                                    numero_accreditation, nbr_pages_annexe, 
    #                                                                                    laboratoire, etat_reception)
    #            db.insertion_table_etalonnage_temp_administration(id_campagne_etal, nbr_pt_etalonnage, 
    #                                                                                    n_sap, identification_instrum, date_etal, 
    #                                                                                    nom_procedure, num_doc, edit_ce, edit_cv, 
    #                                                                                    technicien, ville_etal, domaine_accreditation, 
    #                                                                                    numero_accreditation, nbr_pages_annexe, 
    #                                                                                    laboratoire)
    #          
    #            #Recuperation de l'id_etalonnage_temp_administration :
    #            id_etalonnage_temp_administration = db.recherche_id_etalonnage_admin()
                
            ##################################################################################
            #Remplissage table etalonnage mesure
            
                #champs : CORRECTION, NUM_DOCUMENT, IDENTIFICATION, NBR_PT_ETALONNAGE, NUM_PT_ETAL, NBR_MESURE_PAR_PT, NUM_MESURE, MESURE_ETAL_NC, MESURE_ETAL_C, MESURE_INSTRUM
                
                #gestion profondeur immersion            
                immersion = []
                qlabel_prof_immersion = [self.textEdit_prof_imm_instrum_1, self.textEdit_prof_imm_instrum_2, self.textEdit_prof_imm_instrum_3, 
                                                    self.textEdit_prof_imm_instrum_4, self.textEdit_prof_imm_instrum_5, self.textEdit_prof_imm_instrum_6, 
                                                    self.textEdit_prof_imm_instrum_7, self.textEdit_prof_imm_instrum_8, self.textEdit_prof_imm_instrum_9, 
                                                    self.textEdit_prof_imm_instrum_10, self.textEdit_prof_imm_instrum_11, self.textEdit_prof_imm_instrum_12, 
                                                    self.textEdit_prof_imm_instrum_13, self.textEdit_prof_imm_instrum_14, self.textEdit_prof_imm_instrum_15, 
                                                    self.textEdit_prof_imm_instrum_16, self.textEdit_prof_imm_instrum_17, self.textEdit_prof_imm_instrum_18, 
                                                    self.textEdit_prof_imm_instrum_19, self.textEdit_prof_imm_instrum_20 ]
                
                for ele in qlabel_prof_immersion[i].toPlainText().split() : 
                    immersion.append(ele)            
                
    
                #gestion resolutions instrument saisie
                resolution_instrument = []
                qlabel_resolution_instrum = [self.textEdit_resolution_instrum_1, self.textEdit_resolution_instrum_2, self.textEdit_resolution_instrum_3, 
                                                        self.textEdit_resolution_instrum_4, self.textEdit_resolution_instrum_5, self.textEdit_resolution_instrum_6, 
                                                        self.textEdit_resolution_instrum_7, self.textEdit_resolution_instrum_8, self.textEdit_resolution_instrum_9, 
                                                        self.textEdit_resolution_instrum_10, self.textEdit_resolution_instrum_11, self.textEdit_resolution_instrum_12, 
                                                        self.textEdit_resolution_instrum_13, self.textEdit_resolution_instrum_14, self.textEdit_resolution_instrum_15, 
                                                        self.textEdit_resolution_instrum_16, self.textEdit_resolution_instrum_17, self.textEdit_resolution_instrum_18, 
                                                        self.textEdit_resolution_instrum_19, self.textEdit_resolution_instrum_20 ]
                    
                for ele in qlabel_resolution_instrum[i].toPlainText().split() : 
                    resolution_instrument.append(ele.replace(",", "."))  
                
                
                
                j=0
                while j < self.SpinBox_nbr_pt_etal.value() :
                    nbr_de_mesure = donnees[j]["nb_acquisition"]
                    
                    k = 0
                    while k < nbr_de_mesure :
                        
                        correction = "'"+str(donnees[j]["corrections_instrum_"+str(i+1)][k])+"'"
                        num_pt_etalonnage ="'"+str(donnees[j]["pt_etalonnag_n"])+"'"
                        nbr_mesure = "'"+str(donnees[j]["nb_acquisition"])+"'"
                        numero_mesure = "'"+str(k+1)+"'"
                        mesure_etal_non_corri = "'"+str(donnees[j]["mesures_etal_brute"][k])+"'"
                        mesure_etal_corri =  "'"+str(donnees[j]["mesures_etal_corri"][k])+"'"
                        mesure_instrum = "'"+str(donnees[j]["mesures_instrum_"+str(i+1)][k])+"'"
    
                        #Insertion des donnees de mesures dans la table ETALONNAGE_MESURE 
                        
                        db. insertion_table_etalonnage_mesure(id_etalonnage_temp_administration, id_campagne_etal, 
                                                                                correction, num_doc, identification_instrum, nbr_pt_etalonnage,
                                                                                num_pt_etalonnage, nbr_mesure, numero_mesure, mesure_etal_non_corri,
                                                                                mesure_etal_corri, mesure_instrum)
    
                                        
                        k += 1
        #######################################################################################################################
                #Remplissage Table etalonnage resultats:
                    
                    #champs
                                    
                    nom_etalon = "'"+str(self.Combobox_etalon_select_2.itemText(donnees[j]["etalon"]))+"'"
                    etalon.append(str(self.Combobox_etalon_select_2.itemText(donnees[j]["etalon"])))
                    
                    nom_generateur = "'"+str(self.Combobox_generateur_select_2.itemText(donnees[j]["generateur"]))+"'"
                    generateur.append(str(self.Combobox_generateur_select_2.itemText(donnees[j]["generateur"])))
                    
                    autoechauffement = "'"+str(donnees[j]["auto_echauffement"])+"'"
                    
                    fuite_thermique ="'"+str(donnees[j]["fuite_thermique"])+"'" 
                    
                    resolution = "'"+str(donnees[j]["Resolution"])+"'" # Attention il sagit de l'incertitude et non de la reso de l 'appareil
                    valeurs_resolution.append(donnees[j]["Resolution"])
                    
                    stabilite = "'"+str(donnees[j]["STAB"])+"'"
                    
                    eiv = "'"+str(donnees[j]["EIV"])+"'"
                    
                    moyenne_etalon_non_corri = "'"+str(donnees[j]["moyenne_etal_brute"])+"'"
                                    
                    moyenne_etalon_corri =  "'"+str(donnees[j]["moyenne_etal_corri"])+"'"
                    valeurs_etalon_corri.append(donnees[j]["moyenne_etal_corri"])
                    
                    moyenne_instrum = "'"+str(donnees[j]["moyenne_instrum_"+str(i+1)])+"'"
                    valeurs_instrum.append(donnees[j]["moyenne_instrum_"+str(i+1)])
                    
                    moyenne_correction = "'"+str(donnees[j]["moyenne_corrections_instrum_"+str(i+1)])+"'"
                    valeurs_corrections.append(donnees[j]["moyenne_corrections_instrum_"+str(i+1)])
                    
                    U = "'"+str(donnees[j]["U"])+"'"
                    valeurs_U.append(donnees[j]["U"])
                    
                    ecart_type = "'"+str(donnees[j]["ecart_type"])+"'"
                    
                    temp_consigne = donnees[j]["temp_consig"]
                    
                    valeur_immersion = immersion[j]
                    
                    valeur_resolution_instrument = resolution_instrument[j]
    
                    #insertion dans la table:
                    db. insertion_table_etalonnage_resultat(temp_consigne, id_campagne_etal, id_etalonnage_temp_administration,
                                                                        nom_etalon, nom_generateur, autoechauffement , fuite_thermique, 
                                                                        resolution, stabilite, eiv, num_doc, identification_instrum, date_etal, 
                                                                        nbr_pt_etalonnage, num_pt_etalonnage, moyenne_etalon_non_corri, 
                                                                        moyenne_etalon_corri, moyenne_instrum, moyenne_correction, U, 
                                                                        ecart_type, valeur_immersion, valeur_resolution_instrument)
    
                    j += 1
                
    #            resolution =lecture_onglet_configuration(self)            
                rapport_etalonnage["resolution"] =  resolution_instrument[0] # resolution instrument
                
                rapport_etalonnage["nbr_pt_etalonnage"] = self.SpinBox_nbr_pt_etal.value()
                rapport_etalonnage["etalon"] = list(set(etalon)) # set : on crée un set qui permet d'eliminer les doublons
                rapport_etalonnage["generateur"] = list(set(generateur)) # set : on crée un set qui permet d'eliminer les doublons
                rapport_etalonnage["moyenne_etalon_corri"] = valeurs_etalon_corri
                rapport_etalonnage["moyenne_instrum"] = valeurs_instrum
                rapport_etalonnage["moyenne_correction"] =valeurs_corrections
                rapport_etalonnage["U"] = valeurs_U
                rapport_etalonnage["immersion"] = immersion
                
                #Gestion du polynome correction =f(Tlue) : et enregistrement dans BDD            
                if self.SpinBox_nbr_pt_etal.value() < 5:
                    ordre_poly = 1
                    poly = calcul_polynome (valeurs_instrum,valeurs_corrections, ordre_poly)
                    a = poly[0]
                    b = poly[1]
                    c = 0
                else:
                    ordre_poly = 2
                    poly = calcul_polynome (valeurs_instrum,valeurs_corrections, ordre_poly)
                    a = poly[0]
                    b = poly[1]
                    c = poly[2]
                
                date_du_jour = datetime.datetime.today()
                date_du_jour_mise_en_forme = "'{}-{}-{}'".format(date_du_jour.year, date_du_jour.month, date_du_jour.day)
                archivage = "FALSE"
                
                #insertion dans la table polynome et archivage des poly precedents            
                db.archivage_polynome(identification_instrum, num_doc, date_etal)
                db.insertion_table_polynome(identification_instrum,date_du_jour_mise_en_forme, num_doc, \
                                                        date_etal, ordre_poly,a, b, c, archivage)
                
                
    
                
                list_textedit_conformite =  [self.textEdit_resultat_instru_1, self.textEdit_resultat_instru_2, self.textEdit_resultat_instru_3,
                                                        self.textEdit_resultat_instru_4, self.textEdit_resultat_instru_5, self.textEdit_resultat_instru_6, 
                                                        self.textEdit_resultat_instru_7, self.textEdit_resultat_instru_8, self.textEdit_resultat_instru_9, 
                                                        self.textEdit_resultat_instru_10, self.textEdit_resultat_instru_11, self.textEdit_resultat_instru_12, 
                                                        self.textEdit_resultat_instru_13, self.textEdit_resultat_instru_14, self.textEdit_resultat_instru_15, 
                                                        self.textEdit_resultat_instru_16, self.textEdit_resultat_instru_17, self.textEdit_resultat_instru_18,
                                                        self.textEdit_resultat_instru_19, self.textEdit_resultat_instru_20]
                                                    
                                                
                list_combobox_emt = [self.comboBox_ref_emt_1, self.comboBox_ref_emt_2, self.comboBox_ref_emt_3, 
                                            self.comboBox_ref_emt_4, self.comboBox_ref_emt_5, self.comboBox_ref_emt_6, 
                                            self.comboBox_ref_emt_7, self.comboBox_ref_emt_8, self.comboBox_ref_emt_9, 
                                            self.comboBox_ref_emt_10, self.comboBox_ref_emt_11, self.comboBox_ref_emt_12, 
                                            self.comboBox_ref_emt_13, self.comboBox_ref_emt_14, self.comboBox_ref_emt_15, 
                                            self.comboBox_ref_emt_16, self.comboBox_ref_emt_17, self.comboBox_ref_emt_18, 
                                            self.comboBox_ref_emt_19, self.comboBox_ref_emt_20]
                
                #gestion textedit conformite:
                resultat_conformite = []
                for ele in list_textedit_conformite[i].toPlainText().split():                
                    resultat_conformite.append(ele)
                
    #            print("resultat cof {}".format(resultat_conformite))
                if not resultat_conformite:
                     rapport_etalonnage["conformite"] = "Néant"
                else:
                    if resultat_conformite[0] == "Non":
                        rapport_etalonnage["conformite"] = resultat_conformite[0] + " " + resultat_conformite[1]
                    else:
                        rapport_etalonnage["conformite"] = resultat_conformite[0]
                
                #Recuperation nom referentiel            
                nom_referentiel_conformite = list_combobox_emt[i].currentText()
                id_referentiel_conformite = db.recherche_id_ref_emt(nom_referentiel_conformite)
                rapport_etalonnage["referentiel_emt"] = db.recherche_nom_ref_emt(id_referentiel_conformite)
               
                #gestion des CE/CV et la sauvegarde 
                certificat = RapportEtalonnage(edit_ce) #appel à la classe RapportEtalonnage
                certificat.mise_en_forme_ce(rapport_etalonnage, dossier, rapport_etalonnage["n_certificat"])
    
                    #CV :
                edition_cv = gestion_edition_cv(self) #retourne une list avec true ou false en fct s'il faut ou non edité le cv
                
                if edition_cv[i] == True:
                    db.insertion_table_conformite_temp_resultat(rapport_etalonnage, date_etal)
                    constat = RapportEtalonnage(edit_ce)
                    nom_constat ="{}V".format(rapport_etalonnage["n_certificat"])
                    constat.mise_en_forme_cv(rapport_etalonnage, dossier, nom_constat)
     
                
                i += 1
            
            
            #effacement des fichiers temporaire (present dans le dossier AppData) 
            
            path =os.path.abspath("AppData/")
          
            for ele in os.listdir(path):
                path_total = str(path + "/"+str(ele))
                
                if os.path.isfile(path_total): # verification qu'il s'agit bien de fichier
                    os.remove(path_total)
            
            #effacement des onglets et retour onglet config
            effacement_onglet_saisie (self)
            efface_onglet_resultat(self)
            reinitialise_onglet_config(self)
            self.tabWidget.setCurrentIndex(0) #on revient à l'onglet configuration
            self.progressBar.setValue(0)
            self.qlabel_campagne_etal_onglet_conclusion.clear()
            self.qlabel_information_onglet_conclusion.clear()
            self.actionEnregistrer.setEnabled(False)
        else:
         pass   
 
        
    @pyqtSlot()
    def on_textEdit_auto_echauffement_selectionChanged(self):
        """
        Slot documentation goes here.
        """
        try:
            tuple = QInputDialog.getText(self, 
                           self.trUtf8("Autoechauffement"), 
                           self.trUtf8("Merci de saisir la valeur"))
                           
            
    #        print(tuple)
            if tuple[1] == True :
                self.textEdit_auto_echauffement.clear()
                autoechauffement = round(float(tuple[0].replace(",", "."))/numpy.sqrt(3), 12)
                    
                i=0
                while i < self.SpinBox_nbr_pt_etal.value() :
                    self.textEdit_auto_echauffement.append(str(autoechauffement))
                    i+=1
                nvlle_incertitude =  calcul_incertitude(self)
                    #incertitude
                self.textEdit_U.clear()
                for ele in nvlle_incertitude :                   
                    self.textEdit_U.append(str(round(ele, 12)))
            else :
                pass
        except ValueError:
            pass
    ##################################################################
    @pyqtSlot()
    def on_actionQuitter_triggered(self):
        """
        Slot documentation goes here.
        """        
        self.close()
   
    ##################################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_1_activated(self, index):
        """
        Slot documentation goes here.
        """
        gestion_declaration_conf(self, 1)
                       
    ##################################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_3_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 3)    
    
    #################################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_4_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 4)
    
    ################################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_2_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 2)
    
    ###############################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_5_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 5)
    
    ###############################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_6_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 6)
    
    ###############################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_7_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 7)
        
    ##############################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_8_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 8)    
            
    ##############################################################        
    @pyqtSlot(int)
    def on_comboBox_ref_emt_9_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 9)
    
    #############################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_10_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 10)  
            
    #############################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_11_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 11)
        
    #############################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_12_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 12)
    
    ############################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_13_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 13)               
            
    ############################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_14_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 14)
    
    ############################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_15_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 15)
    
    ############################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_16_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 16)
    
    ############################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_17_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 17)
    
    ############################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_18_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 18)
    
    ############################################################    
    @pyqtSlot(int)
    def on_comboBox_ref_emt_19_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 19)
    
    #############################################################
    @pyqtSlot(int)
    def on_comboBox_ref_emt_20_activated(self, index):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        gestion_declaration_conf(self, 20)
        
        
    @pyqtSlot()
    def on_toolButton_clicked(self):
        """
        Boutton qui ouvre un widget avec les valeurs affichées mais misent à la resolution de l'instrument
        """
        nom_instrum = self.qlabel_nom_instrument_onglet_resultats.text()
        resolution_onglet_config = [self.textEdit_resolution_instrum_1,self.textEdit_resolution_instrum_2, self.textEdit_resolution_instrum_3, 
                                            self.textEdit_resolution_instrum_4, self.textEdit_resolution_instrum_5, self.textEdit_resolution_instrum_6, 
                                            self.textEdit_resolution_instrum_7, self.textEdit_resolution_instrum_8, self.textEdit_resolution_instrum_9, 
                                            self.textEdit_resolution_instrum_10, self.textEdit_resolution_instrum_11, self.textEdit_resolution_instrum_12, 
                                            self.textEdit_resolution_instrum_13, self.textEdit_resolution_instrum_14, self.textEdit_resolution_instrum_15, 
                                            self.textEdit_resolution_instrum_16, self.textEdit_resolution_instrum_17, self.textEdit_resolution_instrum_18, 
                                            self.textEdit_resolution_instrum_19, self.textEdit_resolution_instrum_20]
        
        donnees = {}
        moyennes_etalon_non_corrigees = []
        moyennes_etalon_corrigees = []
        moyennes_instrument = []
        corrections = []
        incertitudes = []
        resolution = []
        n_instrument = int(self.label_5_inst_n.text())
        
        
        for ele in resolution_onglet_config[n_instrument-1].toPlainText().split():
            resolution.append(str(ele.replace(",", ".")))
            
        i=0    
        for ele in self.textEdit_moyennes_etal_brutes.toPlainText().split():
            conversion_ele = str(ele)
            valeur = decimal.Decimal(conversion_ele).quantize(decimal.Decimal(resolution[i]), rounding=decimal.ROUND_HALF_EVEN)
            moyennes_etalon_non_corrigees.append(valeur)  
            
            i+=1
        j=0    
        for ele in self.textEdit_moyennes_etal_corriges.toPlainText().split():
            conversion_ele = str(ele)
            valeur = decimal.Decimal(conversion_ele).quantize(decimal.Decimal(resolution[j]), rounding=decimal.ROUND_HALF_EVEN)
            moyennes_etalon_corrigees.append(valeur)
            
            j+=1
         
        k=0
        for ele in self.textEdit_moyennes_inst.toPlainText().split():
            conversion_ele = str(ele)
            valeur = decimal.Decimal(conversion_ele).quantize(decimal.Decimal(resolution[k]), rounding=decimal.ROUND_HALF_EVEN)
            moyennes_instrument.append(valeur)
            k+=1
        l=0    
        for ele in self.textEdit_corrections.toPlainText().split():
            conversion_ele = str(ele)
            valeur = decimal.Decimal(conversion_ele).quantize(decimal.Decimal(resolution[l]), rounding=decimal.ROUND_HALF_EVEN)
            corrections.append(valeur)
            l+=1
        m=0
        for ele in self.textEdit_U.toPlainText().split():
            conversion_ele = str(ele)
            valeur = decimal.Decimal(conversion_ele).quantize(decimal.Decimal(resolution[m]), rounding=decimal.ROUND_UP)
            incertitudes.append(valeur)
            m+=1
            
        donnees["moyennes_etalon_non_corrigees"] = moyennes_etalon_non_corrigees
        donnees["moyennes_etalon_corrigees"] = moyennes_etalon_corrigees
        donnees["moyennes_instrument"] = moyennes_instrument
        donnees["corrections"] = corrections
        donnees["incertitudes"] = incertitudes
        
        #ouverture widget
        self.resume = Form(nom_instrum, self.SpinBox_nbr_pt_etal.value(),donnees )
        self.resume.setWindowModality(QtCore.Qt.ApplicationModal)
        self.resume.show()
        
########################################################################################################################        
        
def effacement_onglet_saisie (self) :
    '''Fonction permettant d'effacer les textbox de l onglet saisie'''
    list_textedit_mesures_instruments = [self.textEdit_mesures_inst_1, self.textEdit_mesures_inst_2,  self.textEdit_mesures_inst_3, 
                                self.textEdit_mesures_inst_4,  self.textEdit_mesures_inst_5,  self.textEdit_mesures_inst_6, 
                                self.textEdit_mesures_inst_7,  self.textEdit_mesures_inst_8,  self.textEdit_mesures_inst_9, 
                                self.textEdit_mesures_inst_10,  self.textEdit_mesures_inst_11,  self.textEdit_mesures_inst_12, 
                                self.textEdit_mesures_inst_13,  self.textEdit_mesures_inst_14,  self.textEdit_mesures_inst_15, 
                                self.textEdit_mesures_inst_16,  self.textEdit_mesures_inst_17,  self.textEdit_mesures_inst_18, 
                                self.textEdit_mesures_inst_19,  self.textEdit_mesures_inst_20 ]
                                
    #effacement des donnees presentes
    i=0
    while i < self.SpinBox_nbr_instruments.value() :
        list_textedit_mesures_instruments[i].clear()
        i+=1
    self.Zone_text_temp_consigne_2.clear()
    self.Zone_texte_chemin_fichier_etalon_2.clear()
    self.textEdit_mesures_etalon_brutes.clear()
    self.textEdit_mesures_etalon_corrigees.clear()

def reaffectation_donnees_onglet_saisie(self, onglet):
    
    try:
        list_textedit_instruments = [self.textEdit_mesures_inst_1, self.textEdit_mesures_inst_2, self.textEdit_mesures_inst_3, 
                                    self.textEdit_mesures_inst_4, self.textEdit_mesures_inst_5, self.textEdit_mesures_inst_6, 
                                    self.textEdit_mesures_inst_7, self.textEdit_mesures_inst_8, self.textEdit_mesures_inst_9, 
                                    self.textEdit_mesures_inst_10, self.textEdit_mesures_inst_11, self.textEdit_mesures_inst_12, 
                                    self.textEdit_mesures_inst_13, self.textEdit_mesures_inst_14, self.textEdit_mesures_inst_15, 
                                    self.textEdit_mesures_inst_16, self.textEdit_mesures_inst_17, self.textEdit_mesures_inst_18, 
                                    self.textEdit_mesures_inst_19, self.textEdit_mesures_inst_20]
        
        list_indice_instrum = ["mesures_inst_1","mesures_inst_2","mesures_inst_3","mesures_inst_4", 
                            "mesures_inst_5","mesures_inst_6","mesures_inst_7","mesures_inst_8", 
                            "mesures_inst_9","mesures_inst_10","mesures_inst_11","mesures_inst_12", 
                            "mesures_inst_13","mesures_inst_14","mesures_inst_15","mesures_inst_16", 
                            "mesures_inst_17","mesures_inst_18","mesures_inst_19","mesures_inst_20"]
        
        self.Zone_text_temp_consigne_2.setText(str(onglet["temp_consig"] ))
        self.Zone_texte_chemin_fichier_etalon_2.setText(onglet["chemin_fichier_etalon"])
        self.Combobox_etalon_select_2.setCurrentIndex(onglet["etalon"])
        self.Combobox_operateur_select_2.setCurrentIndex(onglet["operateur"])
        self.Combobox_generateur_select_2.setCurrentIndex(onglet["generateur"])
           
        
        for element in onglet["mesures_etal_brute"] :
            self.textEdit_mesures_etalon_brutes.append(str(element))
        for element in onglet["mesures_etal_corri"] :
            self.textEdit_mesures_etalon_corrigees.append(str(round(element, 4)))
        i = 0
        while i < self.SpinBox_nbr_instruments.value() :
            for element in onglet[list_indice_instrum[i]] :
                list_textedit_instruments[i].append(str(element))
            i+=1
    except KeyError:
        pass
####################################################################################################################################

def traitement_donnees_saisies(self):
    '''Traitement des donnees precedentes dans l'onglet saisie :  
    va chercher le poly de correction etalon dans la bdd.
    corrige les mesures calcul moyenne ecart type etc renvoie un dictionnaire avec ces valeurs'''
           
    try :
        
        list_textedit_instruments = [self.textEdit_mesures_inst_1, self.textEdit_mesures_inst_2, self.textEdit_mesures_inst_3, 
                                            self.textEdit_mesures_inst_4, self.textEdit_mesures_inst_5, self.textEdit_mesures_inst_6, 
                                            self.textEdit_mesures_inst_7, self.textEdit_mesures_inst_8, self.textEdit_mesures_inst_9, 
                                            self.textEdit_mesures_inst_10, self.textEdit_mesures_inst_11, self.textEdit_mesures_inst_12, 
                                            self.textEdit_mesures_inst_13, self.textEdit_mesures_inst_14, self.textEdit_mesures_inst_15, 
                                            self.textEdit_mesures_inst_16, self.textEdit_mesures_inst_17, self.textEdit_mesures_inst_18, 
                                            self.textEdit_mesures_inst_19, self.textEdit_mesures_inst_20]

    
        #Mise en forme des donnees instruments
        mesures_instruments= []
        i=0
        while i< self.SpinBox_nbr_instruments.value() :
            mesures_instruments.append(mise_enforme_saisies(list_textedit_instruments[i]))
            i+=1
        
           #Mise en forme donnees etalon brutes :
        
        mesures_brutes = mise_enforme_saisies(self.textEdit_mesures_etalon_brutes)
        
        #Acces bdd
        db = GestionBdd('db')
        db.reconnexion()
        
        #correction des donnees etalons
              
        ident_etalon = str(self.Combobox_etalon_select_2.currentText())
        coeff_poly = db.recuperation_poly_etalon_bis(ident_etalon )
        a = coeff_poly[0]
        b = coeff_poly[1]
        c = coeff_poly[2]

    #        Traitement des donnees brutes de l'etalon 
        i=0
        mesures_corrigees = []
        nbr_mesures = self.SpinBox_nbr_releves_2.value()
        self.textEdit_mesures_etalon_corrigees.clear()
        
                
        while i < nbr_mesures :
            
            # Mise en forme des donnees : 4 chiffres apres la virgule
                      
            corrige = numpy.float64(mesures_brutes[i] +(a*mesures_brutes[i]*mesures_brutes[i]+b*mesures_brutes[i]+c))
            mesures_corrigees.append(corrige)
            self.textEdit_mesures_etalon_corrigees.append(str(round(corrige, 4)))
            
            i +=1
                
        
        #Copie des donnees de l'onglet dans un dictionnaire avant une sauvegarde en picker
        
        donnees_saisies = {}
        #remplissage
        donnees_saisies["temp_consig"] = int(self.Zone_text_temp_consigne_2.toPlainText())
        if self.Zone_text_temp_consigne_2.toPlainText() == "": 
            raise ValueError("pas de valeur de consigne")
                
        donnees_saisies["operateur"] = self.Combobox_operateur_select_2.currentIndex ()
        donnees_saisies["generateur"] = self.Combobox_generateur_select_2.currentIndex ()
        donnees_saisies["etalon"] = self.Combobox_etalon_select_2.currentIndex ()
        donnees_saisies["chemin_fichier_etalon"] = self.Zone_texte_chemin_fichier_etalon_2.toPlainText()
        donnees_saisies["pt_etalonnag_n"] = self.lineedit_pt_etal_n_2.text()
        donnees_saisies["nb_acquisition"] = self.SpinBox_nbr_releves_2.value()
        donnees_saisies["mesures_etal_brute"] = mesures_brutes
        donnees_saisies["mesures_etal_corri"] = mesures_corrigees
        
        list_indice_instrum = ["mesures_inst_1","mesures_inst_2","mesures_inst_3","mesures_inst_4", 
                                        "mesures_inst_5","mesures_inst_6","mesures_inst_7","mesures_inst_8", 
                                        "mesures_inst_9","mesures_inst_10","mesures_inst_11","mesures_inst_12", 
                                        "mesures_inst_13","mesures_inst_14","mesures_inst_15","mesures_inst_16", 
                                        "mesures_inst_17","mesures_inst_18","mesures_inst_19","mesures_inst_20"]
                                        
        
        
        i=0
        while i< self.SpinBox_nbr_instruments.value() :
            
            donnees_saisies[list_indice_instrum[i]] = mesures_instruments[i]
            
            #calcul sur les saisies 
                #moyennes :
            donnees_saisies["moyenne_etal_corri"]  = numpy.mean(donnees_saisies["mesures_etal_corri"],dtype=numpy.float64) #round(numpy.sum(donnees_saisies["mesures_etal_corri"])/self.SpinBox_nbr_releves_2.value(), 4)
            
            donnees_saisies["moyenne_etal_brute"]  = numpy.mean(donnees_saisies["mesures_etal_brute"],dtype=numpy.float64) #round(numpy.sum(donnees_saisies["mesures_etal_brute"])/self.SpinBox_nbr_releves_2.value(), 4)
            donnees_saisies["moyenne_instrum"+"_"+str(i+1)]  = numpy.mean(donnees_saisies[list_indice_instrum[i]],dtype=numpy.float64)#round(numpy.sum(donnees_saisies[list_indice_instrum[i]])/self.SpinBox_nbr_releves_2.value(), 4)
            
                #  corrections
            
            corrections_instruments = []
            
            vector1 = numpy.array(mesures_corrigees)
            vector2 = numpy.array(numpy.float64(mesures_instruments[i]))
            corrections_instruments = numpy.subtract( vector1 ,  vector2)
                        
            donnees_saisies["corrections_instrum"+"_"+str(i+1)] = corrections_instruments
            donnees_saisies["moyenne_corrections_instrum"+"_"+str(i+1)] = numpy.mean(corrections_instruments, dtype=numpy.float64)#round((numpy.sum(corrections_instruments)/self.SpinBox_nbr_releves_2.value()), 4)
            vector3 =  numpy.array(donnees_saisies["corrections_instrum"+"_"+str(i+1)])
            donnees_saisies["ecartype_corrections_instrum"+"_"+str(i+1)] = numpy.std(vector3, dtype=numpy.float64, ddof=1)
            #gestion des futures calcul fuite_thermique et autoechaffement
            donnees_saisies["fuite_thermique_instrum"+"_"+str(i+1)] = "0"
            donnees_saisies["auto_echauffement_instrum"+"_"+str(i+1)] = "0"
            
            
            i+=1
            
        return donnees_saisies
            
    except TypeError:
        pass
    except ValueError:
        QMessageBox.critical(self, 
                                            self.trUtf8("Attention"),
                                            self.trUtf8("Vous n'avez pas saisie de temperature de consigne"))
        
##############################################################################################################        

def chargement_donnees_instrument(self, num_instrum):
    '''Fonction qui recupere l'ensemble des fichiers (pickler) d'un instrument dans une list et la renvoie'''
    num = str(num_instrum)
    num_instrum = []
    
    i=0
    while i < self.SpinBox_nbr_pt_etal.value() :
        nom_fichier = str(i+1) + "_instrum_"+ num
        num_instrum.append(lecture_onglet_saisie(nom_fichier))

        i+=1

    return num_instrum 

###############################################################################################################################

def efface_onglet_resultat(self):
    qlabel = [self.textEdit_pt_consignes, self.textEdit_generateurs, self.textEdit_etalons, 
                self.textEdit_moyennes_etal_brutes, self.textEdit_moyennes_etal_corriges, 
                self.textEdit_moyennes_inst, self.textEdit_corrections, self.textEdit_ecarts_types, 
                self.textEdit_eiv_caract, self.textEdit_stabilite, self.textEdit_resolution, self.textEdit_U,  self.textEdit_fuite_thermique, self.textEdit_auto_echauffement]
    
    for elem in qlabel :
        elem.clear()
    
##############################################################################################################################    
    
def affichage_onglet_resultat (self, donnees, num_instrument) :
    '''fonction affichant les donnees d'un instrument sur l'onglet resultat'''
    
    self.tab_3.setEnabled(True)
    
    i=0
    
    while i < self.SpinBox_nbr_pt_etal.value() :
        
        #text box temp consigne
        self.textEdit_pt_consignes.append(str(donnees[i]["temp_consig"])) 
        #generateur
        if donnees[i]["generateur"] == 0 :
            generateur = "BGF"
        elif donnees[i]["generateur"] == 1 :
            generateur = "HART Scientifique N°1"
        elif donnees[i]["generateur"] == 2 :
            generateur = "HART Scientifique N°2"
        elif donnees[i]["generateur"] == 3 :
            generateur = "Etuve ESPEC  N° 1"
        else :
            generateur ="Etuve ESPEC  N° 2"    
        self.textEdit_generateurs.append(generateur)
        
        
        
        #etalons
        if donnees[i]["etalon"] == 0 :
            etalon = "THE REF 003"
        elif donnees[i]["etalon"] == 1 :
            etalon = "THE REF 004"
        elif donnees[i]["etalon"] == 2 :
            etalon = "THE REF 005"
        elif donnees[i]["etalon"] == 3 :
            etalon = "THE REF 006"
        elif donnees[i]["etalon"] == 4 :
            etalon = "THE REF 007"
        else :
            etalon ="THE REF 008"    
       
       #recuperation donnees EIV meilleure ou courantes
        eiv =  recuperation_EIV_meilleures_ou_courantes(self, generateur, etalon)
        
        stab = calcul_stabilite(self, donnees, num_instrument)
        
        self.textEdit_etalons.append(etalon)
        #moyenne etal brut
        self.textEdit_moyennes_etal_brutes.append(str(round(donnees[i]["moyenne_etal_brute"], 4)))
        #moyenne etal corrige
        self.textEdit_moyennes_etal_corriges.append(str(round(donnees[i]["moyenne_etal_corri"], 4)))
        #moyenne de l'instrument
        self.textEdit_moyennes_inst.append(str(round(donnees[i]["moyenne_instrum"+"_"+str(num_instrument)], 4)))
        #moyenne des corrections
        self.textEdit_corrections.append(str(round(donnees[i]["moyenne_corrections_instrum"+"_"+str(num_instrument)], 4)))
        #ecart type sur les corrections
        ecart_type_corrections = donnees[i]["ecartype_corrections_instrum"+"_"+str(num_instrument)]
        self.textEdit_ecarts_types.append(str(round(ecart_type_corrections, 12)))
        #EIV caracterisation
        self.textEdit_eiv_caract.append(str(eiv))
        #Stabilité
        self.textEdit_stabilite.append(str(stab))
        #fuite thermique
        self.textEdit_fuite_thermique.append(str(donnees[i]["fuite_thermique_instrum"+"_"+str(num_instrument)]))
        #Autoechauffement
        
#        print(donnees[i]["auto_echauffement_instrum_"+str(num_instrument)])
        self.textEdit_auto_echauffement.append(str(donnees[i]["auto_echauffement_instrum_"+str(num_instrument)]))
      
        

        i+=1
       
    #Resolution
    resolution_onglet_config = [self.textEdit_resolution_instrum_1,self.textEdit_resolution_instrum_2, self.textEdit_resolution_instrum_3, 
                                            self.textEdit_resolution_instrum_4, self.textEdit_resolution_instrum_5, self.textEdit_resolution_instrum_6, 
                                            self.textEdit_resolution_instrum_7, self.textEdit_resolution_instrum_8, self.textEdit_resolution_instrum_9, 
                                            self.textEdit_resolution_instrum_10, self.textEdit_resolution_instrum_11, self.textEdit_resolution_instrum_12, 
                                            self.textEdit_resolution_instrum_13, self.textEdit_resolution_instrum_14, self.textEdit_resolution_instrum_15, 
                                            self.textEdit_resolution_instrum_16, self.textEdit_resolution_instrum_17, self.textEdit_resolution_instrum_18, 
                                            self.textEdit_resolution_instrum_19, self.textEdit_resolution_instrum_20]
    
    textedit =resolution_onglet_config[num_instrument - 1]
    self.textEdit_resolution.clear()
#    i=0
#    while i < self.SpinBox_nbr_pt_etal.value() :
    
    for ele in textedit.toPlainText().split() :
        valeur_resolution = float(ele.replace(",", "."))
        incertitude_reso = str(round((valeur_resolution / (2*numpy.sqrt(3))), 12))
        self.textEdit_resolution.append(incertitude_reso)
        
#        i +=1
    
    #incertitude
    for ele in calcul_incertitude(self) :
        self.textEdit_U.append(str(round(ele, 12)))
    
#######################################################################################################################################################

def recuperation_EIV_meilleures_ou_courantes(self, nom_generateur, nom_etalon):
    '''fonction permettant de recupere les donnees d'incertitudes dues aux generateurs et etalons'''    
  
    #Acces bdd
    db = GestionBdd('db')
    db.reconnexion()
    u_gen = db. recuperation_eiv(nom_generateur, nom_etalon)
    
    return u_gen

#########################################################################################################
    
       
      

       
        
def calcul_stabilite(self, donnees, num_instrument):
    '''permet de calculer la stabilité de l'instrument en faisant la difference entre les corrections à 0°C'''

    i=0
    stab = []
    while i < self.SpinBox_nbr_pt_etal.value() :   
        
        if donnees[i]["generateur"] == 0 :
            stab.append(donnees[i]["moyenne_corrections_instrum"+"_"+str(num_instrument)])
            
         
        elif donnees[i]["generateur"] == 3 or donnees[i]["generateur"] == 4 :
            stab.append(donnees[0]["moyenne_corrections_instrum"+"_"+str(num_instrument)])
            stab.append(donnees[(self.SpinBox_nbr_pt_etal.value()-1)]["moyenne_corrections_instrum"+"_"+str(num_instrument)])
        i+=1
    if len(stab) < 1:
        delta = 0
    else:
        delta = round((numpy.amax(stab) - numpy.amin(stab))/(2*numpy.sqrt(3)), 12)
 
    return delta

################################################################################################################################"
        
def calcul_incertitude(self) :
       
        qlabel = [ self.textEdit_ecarts_types.toPlainText(), self.textEdit_eiv_caract.toPlainText(), self.textEdit_stabilite.toPlainText(), 
                    self.textEdit_resolution.toPlainText(), self.textEdit_fuite_thermique.toPlainText(),self.textEdit_auto_echauffement.toPlainText()  ]
       
        i = 0
        j = 0
        
        nbr_pt_temp = self.SpinBox_nbr_pt_etal.value()
        nbr_indice_qlabel = 5  
        valeur_qlabel =[]
        
        while i <= nbr_indice_qlabel :
            for ele in qlabel[i].split() :
                
                valeur_qlabel.append(ele)
            i+=1
                
        
        incertitude = []
               
        while j < nbr_pt_temp :
            
            
            incertitude.append(2 *numpy.sqrt( float(valeur_qlabel[j]) * float(valeur_qlabel[j])
                    + float(valeur_qlabel[j+nbr_pt_temp]) * float(valeur_qlabel[j+nbr_pt_temp])
                    + float(valeur_qlabel[j+2*(nbr_pt_temp)]) * float(valeur_qlabel[j+2*(nbr_pt_temp)])
                    + float(valeur_qlabel[j+3*(nbr_pt_temp)])  * float(valeur_qlabel[j+3*(nbr_pt_temp)]) 
                    + float(valeur_qlabel[j+4*(nbr_pt_temp)]) * float(valeur_qlabel[j+4*(nbr_pt_temp)])
                    + float(valeur_qlabel[j+5*(nbr_pt_temp)])  * float(valeur_qlabel[j+5*(nbr_pt_temp)])))    
            
            
            j+=1
        
        return incertitude

######################################################################################################################

def onglet_resultat(self, num_instrument):
            #chargement des donnees depuis repertoir AppData
    
    result_instrument = chargement_donnees_instrument(self, num_instrument)
    

    efface_onglet_resultat(self)
    affichage_onglet_resultat (self, result_instrument, num_instrument)
    self.label_5_inst_n.setText(str(num_instrument))
    self.label_7_nb_inst.setText(str(self.SpinBox_nbr_instruments.value()))
    
#####################################################################################################        
      
def sauvegarde_onglet_resultat(self, num_instrument):
    '''fonction qui sauvegarde l'onglet affiché et utilise la fonction sauvegarde_onglet_saisie pour mettre ds un pickler'''
    
    result_instrument = chargement_donnees_instrument(self, num_instrument)
#    print(result_instrument)
    
    U = []
    for ele in self.textEdit_U.toPlainText().split() :
        U.append(ele)
    
    ecart_type = []
    for ele in self.textEdit_ecarts_types.toPlainText().split() :
        ecart_type.append(ele)
    
    eiv_caract = []
    for ele in self.textEdit_eiv_caract.toPlainText().split() :
        eiv_caract.append(ele)
    
    stab = []
    for ele in self.textEdit_stabilite.toPlainText().split():
        stab.append(ele)
    
    resolution = []
    for ele in self.textEdit_resolution.toPlainText().split() :
        resolution.append(ele)
    
    fuite_thermique = []
    for ele in  self.textEdit_fuite_thermique.toPlainText().split() :
        fuite_thermique.append(ele)
        
    autoechauffement = []
    for ele in  self.textEdit_auto_echauffement.toPlainText().split() :
        autoechauffement.append(ele)
    
    onglet_resutat = []
    i = 0
    while i < self.SpinBox_nbr_pt_etal.value():
        dictionnaire = result_instrument[i]
        dictionnaire["U"] = U[i]
        dictionnaire["ecart_type"] = ecart_type[i]        
        dictionnaire["EIV"] = eiv_caract[i]
        dictionnaire["STAB"] = stab[i]
        dictionnaire["Resolution"] = resolution[i]
        dictionnaire["fuite_thermique"] = fuite_thermique[i]
        dictionnaire["auto_echauffement"] = autoechauffement[i]
        
    #on modifie le fichier n°tem_instrum_n° pour reaffecter la fuite thermique et auto echauf avec les valeurs saisies dans les textbox 

        nom_fichier_bis = str(i+1)+"_instrum_"+ str(num_instrument)
        donnees_a_modifier = lecture_onglet_saisie(nom_fichier_bis)
        donnees_a_modifier["fuite_thermique_instrum_"+str(num_instrument)] = fuite_thermique[i]
        donnees_a_modifier["auto_echauffement_instrum_"+str(num_instrument)] = autoechauffement[i]
        
        sauvegarde_onglet_saisie(donnees_a_modifier , nom_fichier_bis)
        
        
        onglet_resutat.append(dictionnaire)
        
        i+=1
        nom_fichier = str("resultat_" + "_instrum_"+ str(num_instrument))
        sauvegarde_onglet_saisie(onglet_resutat , nom_fichier)
        
        
           
#########################################################################################################
        
def numero_certificat(self):
    '''Fonction permettant de regarder dans la base le numero d'etalonnage du certifcat en regardant tous les talonnage fait dans l'annee'''
        
    date = self.calendarWidget.selectedDate()
    annee = date.toString('yyyy')
    mois = date.toString('MM')
#    annee =date.year()
    #Acces bdd
    db = GestionBdd('db')
    db.reconnexion()

    results = db.n_certificat_3(annee)
    
    id_etal = results[1]
    numero = str(len(results[0])+1)
  
    numero_certificat ="T{}{}_{}".format(annee, mois, numero)
    return numero_certificat, id_etal
    
#############################################################################################################################################################################################
def numero_campagne_etal(self):
    '''fonction equivalente à numero certificat mais pour la campagne d'etalonnage'''
    date = self.calendarWidget.selectedDate()
    annee =date.year()#Acces bdd
    
    db = GestionBdd('db')
    db.reconnexion()
    
    results = db.n_campagne_etal_3(annee)
    
    id_campagne = results[1]
    numero_campagne_etal =str(str(len(results[0])+1)+"_"+str(date.year()))
    
    return numero_campagne_etal, id_campagne

####################################################################
def lecture_onglet_configuration(self) :
    '''fonction permettant de lire l'onglet configuration chaque instrument configuré est mis dans un dictionnaire et on met le tout dans une list
    La fonction retourne la list avec l'ensemble des parametres saisies dans l'onglet config'''
    
    qlabel = [self.comboBox_ident_instrum_1, self.textEdit_n_serie_instrum_1, self.textEdit_construct_instrum_1, self.textEdit_type_instrum_1, self.textEdit_resolution_instrum_1, self.textEdit_prof_imm_instrum_1, self.radioButton_cofrac_1, self.radioButton_non_cofrac_1, self.textEdit_com_inst_1, 
                self.comboBox_ident_instrum_2, self.textEdit_n_serie_instrum_2, self.textEdit_construct_instrum_2, self.textEdit_type_instrum_2, self.textEdit_resolution_instrum_2, self.textEdit_prof_imm_instrum_2, self.radioButton_cofrac_2, self.radioButton_non_cofrac_2, self.textEdit_com_inst_2,
                self.comboBox_ident_instrum_3, self.textEdit_n_serie_instrum_3, self.textEdit_construct_instrum_3, self.textEdit_type_instrum_3, self.textEdit_resolution_instrum_3, self.textEdit_prof_imm_instrum_3, self.radioButton_cofrac_3, self.radioButton_non_cofrac_3, self.textEdit_com_inst_3,
                self.comboBox_ident_instrum_4, self.textEdit_n_serie_instrum_4, self.textEdit_construct_instrum_4, self.textEdit_type_instrum_4, self.textEdit_resolution_instrum_4, self.textEdit_prof_imm_instrum_4, self.radioButton_cofrac_4, self.radioButton_non_cofrac_4, self.textEdit_com_inst_4,
                self.comboBox_ident_instrum_5, self.textEdit_n_serie_instrum_5, self.textEdit_construct_instrum_5, self.textEdit_type_instrum_5, self.textEdit_resolution_instrum_5, self.textEdit_prof_imm_instrum_5, self.radioButton_cofrac_5, self.radioButton_non_cofrac_5, self.textEdit_com_inst_5,
                self.comboBox_ident_instrum_6, self.textEdit_n_serie_instrum_6, self.textEdit_construct_instrum_6, self.textEdit_type_instrum_6, self.textEdit_resolution_instrum_6, self.textEdit_prof_imm_instrum_6, self.radioButton_cofrac_6, self.radioButton_non_cofrac_6, self.textEdit_com_inst_6,
                self.comboBox_ident_instrum_7, self.textEdit_n_serie_instrum_7, self.textEdit_construct_instrum_7, self.textEdit_type_instrum_7, self.textEdit_resolution_instrum_7, self.textEdit_prof_imm_instrum_7, self.radioButton_cofrac_7, self.radioButton_non_cofrac_7, self.textEdit_com_inst_7,
                self.comboBox_ident_instrum_8, self.textEdit_n_serie_instrum_8, self.textEdit_construct_instrum_8, self.textEdit_type_instrum_8, self.textEdit_resolution_instrum_8, self.textEdit_prof_imm_instrum_8, self.radioButton_cofrac_8, self.radioButton_non_cofrac_8, self.textEdit_com_inst_8,
                self.comboBox_ident_instrum_9, self.textEdit_n_serie_instrum_9, self.textEdit_construct_instrum_9, self.textEdit_type_instrum_9, self.textEdit_resolution_instrum_9, self.textEdit_prof_imm_instrum_9, self.radioButton_cofrac_9, self.radioButton_non_cofrac_9, self.textEdit_com_inst_9,
                self.comboBox_ident_instrum_10, self.textEdit_n_serie_instrum_10, self.textEdit_construct_instrum_10, self.textEdit_type_instrum_10, self.textEdit_resolution_instrum_10, self.textEdit_prof_imm_instrum_10, self.radioButton_cofrac_10, self.radioButton_non_cofrac_10, self.textEdit_com_inst_10,
                self.comboBox_ident_instrum_11, self.textEdit_n_serie_instrum_11, self.textEdit_construct_instrum_11, self.textEdit_type_instrum_11, self.textEdit_resolution_instrum_11, self.textEdit_prof_imm_instrum_11, self.radioButton_cofrac_11, self.radioButton_non_cofrac_11, self.textEdit_com_inst_11,
                self.comboBox_ident_instrum_12, self.textEdit_n_serie_instrum_12, self.textEdit_construct_instrum_12, self.textEdit_type_instrum_12, self.textEdit_resolution_instrum_12, self.textEdit_prof_imm_instrum_12, self.radioButton_cofrac_12, self.radioButton_non_cofrac_12, self.textEdit_com_inst_12,
                self.comboBox_ident_instrum_13, self.textEdit_n_serie_instrum_13, self.textEdit_construct_instrum_13, self.textEdit_type_instrum_13, self.textEdit_resolution_instrum_13, self.textEdit_prof_imm_instrum_13, self.radioButton_cofrac_13, self.radioButton_non_cofrac_13, self.textEdit_com_inst_13,
                self.comboBox_ident_instrum_14, self.textEdit_n_serie_instrum_14, self.textEdit_construct_instrum_14, self.textEdit_type_instrum_14, self.textEdit_resolution_instrum_14, self.textEdit_prof_imm_instrum_14, self.radioButton_cofrac_14, self.radioButton_non_cofrac_14, self.textEdit_com_inst_14,
                self.comboBox_ident_instrum_15, self.textEdit_n_serie_instrum_15, self.textEdit_construct_instrum_15, self.textEdit_type_instrum_15, self.textEdit_resolution_instrum_15, self.textEdit_prof_imm_instrum_15, self.radioButton_cofrac_15, self.radioButton_non_cofrac_15, self.textEdit_com_inst_15,
                self.comboBox_ident_instrum_16, self.textEdit_n_serie_instrum_16, self.textEdit_construct_instrum_16, self.textEdit_type_instrum_16, self.textEdit_resolution_instrum_16, self.textEdit_prof_imm_instrum_16, self.radioButton_cofrac_16, self.radioButton_non_cofrac_16, self.textEdit_com_inst_16,
                self.comboBox_ident_instrum_17, self.textEdit_n_serie_instrum_17, self.textEdit_construct_instrum_17, self.textEdit_type_instrum_17, self.textEdit_resolution_instrum_17, self.textEdit_prof_imm_instrum_17, self.radioButton_cofrac_17, self.radioButton_non_cofrac_17, self.textEdit_com_inst_17,
                self.comboBox_ident_instrum_18, self.textEdit_n_serie_instrum_18, self.textEdit_construct_instrum_18, self.textEdit_type_instrum_18, self.textEdit_resolution_instrum_18, self.textEdit_prof_imm_instrum_18, self.radioButton_cofrac_18, self.radioButton_non_cofrac_18, self.textEdit_com_inst_18,
                self.comboBox_ident_instrum_19, self.textEdit_n_serie_instrum_19, self.textEdit_construct_instrum_19, self.textEdit_type_instrum_19, self.textEdit_resolution_instrum_19, self.textEdit_prof_imm_instrum_19, self.radioButton_cofrac_19, self.radioButton_non_cofrac_19, self.textEdit_com_inst_19,
                self.comboBox_ident_instrum_20, self.textEdit_n_serie_instrum_20, self.textEdit_construct_instrum_20, self.textEdit_type_instrum_20, self.textEdit_resolution_instrum_20, self.textEdit_prof_imm_instrum_20, self.radioButton_cofrac_20, self.radioButton_non_cofrac_20, self.textEdit_com_inst_20
                ]
    list_text_edit_etat_reception = [self.textEdit_etat_reception_inst_1, self.textEdit_etat_reception_inst_2, self.textEdit_etat_reception_inst_3,
                                                                self.textEdit_etat_reception_inst_4, self.textEdit_etat_reception_inst_5, self.textEdit_etat_reception_inst_6, 
                                                                self.textEdit_etat_reception_inst_7, self.textEdit_etat_reception_inst_8, self.textEdit_etat_reception_inst_9,
                                                                self.textEdit_etat_reception_inst_10, self.textEdit_etat_reception_inst_11, self.textEdit_etat_reception_inst_12,
                                                                self.textEdit_etat_reception_inst_13, self.textEdit_etat_reception_inst_14, self.textEdit_etat_reception_inst_15,
                                                                self.textEdit_etat_reception_inst_16, self.textEdit_etat_reception_inst_17, self.textEdit_etat_reception_inst_18,
                                                                self.textEdit_etat_reception_inst_19, self.textEdit_etat_reception_inst_20]
   
    nbr_instrum = self.SpinBox_nbr_instruments.value()    
    nbr_qlabel = nbr_instrum * 9 # on multiplie par 8 car il ya  8 label par instrum (7 labels et 2 radiobuttons)
    
    onglet_config = []
    dict_def_campagne_etal = {"nbr_instrum": self.SpinBox_nbr_instruments.value(),"nbr_pts_temp": self.SpinBox_nbr_pt_etal.value(), 
                                            "Date": self.calendarWidget.selectedDate() }
    
    onglet_config.append(dict_def_campagne_etal)
    
        
    i = 0
    while i < nbr_qlabel : #c'est le nbr de label qui compte puisque c'est un multiple du nbr d'instrum
        dict_saisie_instrum = {}
        resolution = []
        immersion = []
        dict_saisie_instrum["indice_instrum"] = qlabel[i].currentIndex ()
        dict_saisie_instrum["nom_instrum"] = qlabel[i].currentText ()
        dict_saisie_instrum["n_serie"] = qlabel[i+1].toPlainText ()
        dict_saisie_instrum["constructeur"] = qlabel[i+2].toPlainText ()
        dict_saisie_instrum["Type"] = qlabel[i+3].toPlainText ()
        
        #resolution : (on met en list car il pourrait y en avoir plusieurs)
        for ele in qlabel[i+4].toPlainText ().split() :
            resolution.append(ele.replace(",", "."))
            
        resolu_1 = resolution[0]
         
        if len(resolution) < self.SpinBox_nbr_pt_etal.value():
            compteur = self.SpinBox_nbr_pt_etal.value() - len(resolution)
            j=0
            while j < compteur:
                resolution.append(resolu_1)
                j+=1
            qlabel[i+4].clear() #on reaaffect l'ensemble des resolutions pour le calucl d'incertitude
            for ele in resolution:
                qlabel[i+4].append(ele)
        else:
            pass
        
        dict_saisie_instrum["resolution"] = resolution
        
        #pronfondeur immersion : (on met en list car il pourrait y en avoir plusieurs)
        
        for ele in qlabel[i+5].toPlainText ().split() :
            immersion.append(ele)
        immer_1 = immersion[0]
                  
        if len(immersion) < self.SpinBox_nbr_pt_etal.value():
            compteur_2 = self.SpinBox_nbr_pt_etal.value() - len(immersion)
            k=0
            while k < compteur_2:
                immersion.append(immer_1)
                k+=1
                
        qlabel[i+5].clear() #on reaaffect l'ensemble des immersions
        for ele in immersion:
            qlabel[i+5].append(ele)
        else:
            pass
                
        dict_saisie_instrum["immersion"] = immersion       
        
        
        #Prestation cofrac ou non
        if qlabel[i+6].isChecked() == True :
            dict_saisie_instrum["Type_etalonnage"] = "Cofrac"
        else :
            dict_saisie_instrum["Type_etalonnage"] = "Non Cofrac"
        
        
        #commentaire
        dict_saisie_instrum["renseignement_complementaire"] = qlabel[i+8].toPlainText ()
        onglet_config.append(dict_saisie_instrum)
        
        i += 9
    i=0
    while i < nbr_instrum :
        onglet_config[i+1]["Etat_reception"] = list_text_edit_etat_reception[i].toPlainText()
        
        i+=1
    
    return onglet_config
 #########################################################      

def gestion_caracteristiques_instrument(self, num_instrument):
        '''Fonction qui gere l'affichage des caracteristique de l'instrument apres selection par un combobox (onglet config)'''
        
        indice = num_instrument - 1 #les lists commence à 0
        
        text_qcombobox = [self.comboBox_ident_instrum_1.currentText (), self.comboBox_ident_instrum_2.currentText (), self.comboBox_ident_instrum_3.currentText (), 
                                    self.comboBox_ident_instrum_4.currentText (), self.comboBox_ident_instrum_5.currentText (), self.comboBox_ident_instrum_6.currentText (), 
                                    self.comboBox_ident_instrum_7.currentText (), self.comboBox_ident_instrum_8.currentText (), self.comboBox_ident_instrum_9.currentText (), 
                                    self.comboBox_ident_instrum_10.currentText (), self.comboBox_ident_instrum_11.currentText (), self.comboBox_ident_instrum_12.currentText (),
                                    self.comboBox_ident_instrum_13.currentText () , self.comboBox_ident_instrum_14.currentText (), self.comboBox_ident_instrum_15.currentText (), 
                                    self.comboBox_ident_instrum_16.currentText (), self.comboBox_ident_instrum_17.currentText (), self.comboBox_ident_instrum_18.currentText (), 
                                    self.comboBox_ident_instrum_19.currentText (), self.comboBox_ident_instrum_20.currentText ()
                                    ]
        
        
        
        
        qlabel = [[self.textEdit_n_serie_instrum_1, self.textEdit_construct_instrum_1, self.textEdit_type_instrum_1, self.textEdit_resolution_instrum_1, self.textEdit_com_inst_1], 
                    [self.textEdit_n_serie_instrum_2, self.textEdit_construct_instrum_2, self.textEdit_type_instrum_2, self.textEdit_resolution_instrum_2, self.textEdit_com_inst_2],
                    [self.textEdit_n_serie_instrum_3, self.textEdit_construct_instrum_3, self.textEdit_type_instrum_3, self.textEdit_resolution_instrum_3, self.textEdit_com_inst_3],
                    [self.textEdit_n_serie_instrum_4, self.textEdit_construct_instrum_4, self.textEdit_type_instrum_4, self.textEdit_resolution_instrum_4, self.textEdit_com_inst_4],
                    [self.textEdit_n_serie_instrum_5, self.textEdit_construct_instrum_5, self.textEdit_type_instrum_5, self.textEdit_resolution_instrum_5, self.textEdit_com_inst_5],
                    [self.textEdit_n_serie_instrum_6, self.textEdit_construct_instrum_6, self.textEdit_type_instrum_6, self.textEdit_resolution_instrum_6, self.textEdit_com_inst_6],
                    [self.textEdit_n_serie_instrum_7, self.textEdit_construct_instrum_7, self.textEdit_type_instrum_7, self.textEdit_resolution_instrum_7, self.textEdit_com_inst_7],
                    [self.textEdit_n_serie_instrum_8, self.textEdit_construct_instrum_8, self.textEdit_type_instrum_8, self.textEdit_resolution_instrum_8, self.textEdit_com_inst_8],
                    [self.textEdit_n_serie_instrum_9, self.textEdit_construct_instrum_9, self.textEdit_type_instrum_9, self.textEdit_resolution_instrum_9, self.textEdit_com_inst_9],
                    [self.textEdit_n_serie_instrum_10, self.textEdit_construct_instrum_10, self.textEdit_type_instrum_10, self.textEdit_resolution_instrum_10, self.textEdit_com_inst_10],
                    [self.textEdit_n_serie_instrum_11, self.textEdit_construct_instrum_11, self.textEdit_type_instrum_11, self.textEdit_resolution_instrum_11, self.textEdit_com_inst_11],
                    [self.textEdit_n_serie_instrum_12, self.textEdit_construct_instrum_12, self.textEdit_type_instrum_12, self.textEdit_resolution_instrum_12, self.textEdit_com_inst_12],
                    [self.textEdit_n_serie_instrum_13, self.textEdit_construct_instrum_13, self.textEdit_type_instrum_13, self.textEdit_resolution_instrum_13, self.textEdit_com_inst_13],
                    [self.textEdit_n_serie_instrum_14, self.textEdit_construct_instrum_14, self.textEdit_type_instrum_14, self.textEdit_resolution_instrum_14, self.textEdit_com_inst_14],
                    [self.textEdit_n_serie_instrum_15, self.textEdit_construct_instrum_15, self.textEdit_type_instrum_15, self.textEdit_resolution_instrum_15, self.textEdit_com_inst_15],
                    [self.textEdit_n_serie_instrum_16, self.textEdit_construct_instrum_16, self.textEdit_type_instrum_16, self.textEdit_resolution_instrum_16, self.textEdit_com_inst_16],
                    [self.textEdit_n_serie_instrum_17, self.textEdit_construct_instrum_17, self.textEdit_type_instrum_17, self.textEdit_resolution_instrum_17, self.textEdit_com_inst_17],
                    [self.textEdit_n_serie_instrum_18, self.textEdit_construct_instrum_18, self.textEdit_type_instrum_18, self.textEdit_resolution_instrum_18, self.textEdit_com_inst_18],
                    [self.textEdit_n_serie_instrum_19, self.textEdit_construct_instrum_19, self.textEdit_type_instrum_19, self.textEdit_resolution_instrum_19, self.textEdit_com_inst_19],
                    [self.textEdit_n_serie_instrum_20, self.textEdit_construct_instrum_20, self.textEdit_type_instrum_20, self.textEdit_resolution_instrum_20, self.textEdit_com_inst_20]
                ]
        
        
        for ele in qlabel[indice]:
            ele.clear()

        
        # acces à la BDD
        db = GestionBdd('db')
        db.reconnexion()
        ident_instrum = text_qcombobox[indice]
        caract_inst = db.gestion_remplissage_onglet_config_2(ident_instrum) #il s'agit dun tupple
        
        n_serie = caract_inst[0]
        constructeur = caract_inst[1]
        type = caract_inst[2]
        resolution = str(caract_inst[3])
        commentaire = caract_inst[4]
    #        print(type(commentaire))
        try:
            qlabel[indice][0].append(n_serie)
            qlabel[indice][1].append(constructeur)
            qlabel[indice][2].append(type)
            qlabel[indice][3].append(resolution)            
            qlabel[indice][4].append(commentaire)

        except TypeError : #se declenche qd il n'y a pas de commentaire dans la base (la variable est une QPyNullVariant comment tester?)
            pass

            
            
###############################################################################################

def gestion_declaration_conf(self, n_combobox_emt):
         '''fonction qui permet de gerer l'affichage du resultat de conformité
         elle appel la classe OngletConformite'''
         # acces à la BDD
         db = GestionBdd('db')
         db.reconnexion()
         path = "AppData/resultat__instrum_"+str(n_combobox_emt)
         clef_lists = n_combobox_emt - 1
         nom_fichier_resultat = "resultat__instrum_"+str(n_combobox_emt)
                
         list_combobox_emt = [self.comboBox_ref_emt_1, self.comboBox_ref_emt_2, self.comboBox_ref_emt_3, 
                                        self.comboBox_ref_emt_4, self.comboBox_ref_emt_5, self.comboBox_ref_emt_6, 
                                        self.comboBox_ref_emt_7, self.comboBox_ref_emt_8, self.comboBox_ref_emt_9, 
                                        self.comboBox_ref_emt_10, self.comboBox_ref_emt_11, self.comboBox_ref_emt_12, 
                                        self.comboBox_ref_emt_13, self.comboBox_ref_emt_14, self.comboBox_ref_emt_15, 
                                        self.comboBox_ref_emt_16, self.comboBox_ref_emt_17, self.comboBox_ref_emt_18,
                                        self.comboBox_ref_emt_19, self.comboBox_ref_emt_20]
                                        
         list_textedit_resultat_instrum = [self.textEdit_resultat_instru_1, self.textEdit_resultat_instru_2, self.textEdit_resultat_instru_3,
                                                    self.textEdit_resultat_instru_4, self.textEdit_resultat_instru_5, self.textEdit_resultat_instru_6, 
                                                    self.textEdit_resultat_instru_7, self.textEdit_resultat_instru_8, self.textEdit_resultat_instru_9, 
                                                    self.textEdit_resultat_instru_10, self.textEdit_resultat_instru_11, self.textEdit_resultat_instru_12, 
                                                    self.textEdit_resultat_instru_13, self.textEdit_resultat_instru_14, self.textEdit_resultat_instru_15, 
                                                    self.textEdit_resultat_instru_16, self.textEdit_resultat_instru_17, self.textEdit_resultat_instru_18, 
                                                    self.textEdit_resultat_instru_19, self.textEdit_resultat_instru_20]
         
         
         list_textedit_resolution = [self.textEdit_resolution_instrum_1.toPlainText(), self.textEdit_resolution_instrum_2.toPlainText(), self.textEdit_resolution_instrum_3.toPlainText(), 
                                            self.textEdit_resolution_instrum_4.toPlainText(), self.textEdit_resolution_instrum_5.toPlainText(), self.textEdit_resolution_instrum_6.toPlainText(), 
                                            self.textEdit_resolution_instrum_7.toPlainText(), self.textEdit_resolution_instrum_8.toPlainText(), self.textEdit_resolution_instrum_9.toPlainText(), 
                                            self.textEdit_resolution_instrum_10.toPlainText(), self.textEdit_resolution_instrum_11.toPlainText(), self.textEdit_resolution_instrum_12.toPlainText(), 
                                            self.textEdit_resolution_instrum_13.toPlainText(), self.textEdit_resolution_instrum_14.toPlainText(), self.textEdit_resolution_instrum_15.toPlainText(), 
                                            self.textEdit_resolution_instrum_16.toPlainText(), self.textEdit_resolution_instrum_17.toPlainText(), self.textEdit_resolution_instrum_18.toPlainText(), 
                                            self.textEdit_resolution_instrum_19.toPlainText(), self.textEdit_resolution_instrum_20.toPlainText()]
       
      
         if os.path.isfile(path):
                        
            #declaration des constantes
            nbr_pt_etal = self.SpinBox_nbr_pt_etal.value()
            donnees = lecture_onglet_saisie(nom_fichier_resultat)
            nom_referentiel = str(list_combobox_emt[clef_lists].currentText())
            id_referentiel_conformite = db.recherche_id_ref_emt(nom_referentiel)
#            print("id_referentiel_conformite recherche bdd {}".format(id_referentiel_conformite))
            n_instrument = n_combobox_emt
            
            #on appel la classe qui gere l'onglet resultat
            declaration_conformite = OngletConformite.declaration_conformite(self,n_instrument, nbr_pt_etal,  donnees, id_referentiel_conformite, list_textedit_resolution)
            
            list_textedit_resultat_instrum[clef_lists].clear()
            list_textedit_resultat_instrum[clef_lists].append(declaration_conformite[0])
            list_textedit_resultat_instrum[clef_lists].append(declaration_conformite[1])
###########################################################################################
def gestion_edition_cv(self):
    
   ''' fonction qui regarde les checkbox cochés et cree une list avec true ou false pour chaque instrum
   true si cv doit etre edité fase sinon ''' 
   

   checkbox_onglet_resultat = [self.checkBox, self.checkBox_2, self.checkBox_3, self.checkBox_4, 
                                            self.checkBox_5, self.checkBox_6, self.checkBox_7, self.checkBox_8, 
                                            self.checkBox_9, self.checkBox_10, self.checkBox_11, self.checkBox_12, 
                                            self.checkBox_13, self.checkBox_14, self.checkBox_15, self.checkBox_16, 
                                            self.checkBox_17, self.checkBox_18, self.checkBox_19, self.checkBox_20]
    
   edition_cv = []
    
   for ele in checkbox_onglet_resultat:
        if ele.isChecked() == True:
            edition_cv.append(True)
        else:
            edition_cv.append(False)
   
   return edition_cv
#######################################################################################################
def reinitialise_onglet_config(self):
    '''fonction qui efface l'onglet config'''
   
    qlabel = [self.comboBox_ident_instrum_1, self.textEdit_n_serie_instrum_1, self.textEdit_construct_instrum_1, self.textEdit_type_instrum_1, self.textEdit_resolution_instrum_1, self.textEdit_prof_imm_instrum_1, self.radioButton_cofrac_1, self.radioButton_non_cofrac_1, self.textEdit_com_inst_1, 
                self.comboBox_ident_instrum_2, self.textEdit_n_serie_instrum_2, self.textEdit_construct_instrum_2, self.textEdit_type_instrum_2, self.textEdit_resolution_instrum_2, self.textEdit_prof_imm_instrum_2, self.radioButton_cofrac_2, self.radioButton_non_cofrac_2, self.textEdit_com_inst_2,
                self.comboBox_ident_instrum_3, self.textEdit_n_serie_instrum_3, self.textEdit_construct_instrum_3, self.textEdit_type_instrum_3, self.textEdit_resolution_instrum_3, self.textEdit_prof_imm_instrum_3, self.radioButton_cofrac_3, self.radioButton_non_cofrac_3, self.textEdit_com_inst_3,
                self.comboBox_ident_instrum_4, self.textEdit_n_serie_instrum_4, self.textEdit_construct_instrum_4, self.textEdit_type_instrum_4, self.textEdit_resolution_instrum_4, self.textEdit_prof_imm_instrum_4, self.radioButton_cofrac_4, self.radioButton_non_cofrac_4, self.textEdit_com_inst_4,
                self.comboBox_ident_instrum_5, self.textEdit_n_serie_instrum_5, self.textEdit_construct_instrum_5, self.textEdit_type_instrum_5, self.textEdit_resolution_instrum_5, self.textEdit_prof_imm_instrum_5, self.radioButton_cofrac_5, self.radioButton_non_cofrac_5, self.textEdit_com_inst_5,
                self.comboBox_ident_instrum_6, self.textEdit_n_serie_instrum_6, self.textEdit_construct_instrum_6, self.textEdit_type_instrum_6, self.textEdit_resolution_instrum_6, self.textEdit_prof_imm_instrum_6, self.radioButton_cofrac_6, self.radioButton_non_cofrac_6, self.textEdit_com_inst_6,
                self.comboBox_ident_instrum_7, self.textEdit_n_serie_instrum_7, self.textEdit_construct_instrum_7, self.textEdit_type_instrum_7, self.textEdit_resolution_instrum_7, self.textEdit_prof_imm_instrum_7, self.radioButton_cofrac_7, self.radioButton_non_cofrac_7, self.textEdit_com_inst_7,
                self.comboBox_ident_instrum_8, self.textEdit_n_serie_instrum_8, self.textEdit_construct_instrum_8, self.textEdit_type_instrum_8, self.textEdit_resolution_instrum_8, self.textEdit_prof_imm_instrum_8, self.radioButton_cofrac_8, self.radioButton_non_cofrac_8, self.textEdit_com_inst_8,
                self.comboBox_ident_instrum_9, self.textEdit_n_serie_instrum_9, self.textEdit_construct_instrum_9, self.textEdit_type_instrum_9, self.textEdit_resolution_instrum_9, self.textEdit_prof_imm_instrum_9, self.radioButton_cofrac_9, self.radioButton_non_cofrac_9, self.textEdit_com_inst_9,
                self.comboBox_ident_instrum_10, self.textEdit_n_serie_instrum_10, self.textEdit_construct_instrum_10, self.textEdit_type_instrum_10, self.textEdit_resolution_instrum_10, self.textEdit_prof_imm_instrum_10, self.radioButton_cofrac_10, self.radioButton_non_cofrac_10, self.textEdit_com_inst_10,
                self.comboBox_ident_instrum_11, self.textEdit_n_serie_instrum_11, self.textEdit_construct_instrum_11, self.textEdit_type_instrum_11, self.textEdit_resolution_instrum_11, self.textEdit_prof_imm_instrum_11, self.radioButton_cofrac_11, self.radioButton_non_cofrac_11, self.textEdit_com_inst_11,
                self.comboBox_ident_instrum_12, self.textEdit_n_serie_instrum_12, self.textEdit_construct_instrum_12, self.textEdit_type_instrum_12, self.textEdit_resolution_instrum_12, self.textEdit_prof_imm_instrum_12, self.radioButton_cofrac_12, self.radioButton_non_cofrac_12, self.textEdit_com_inst_12,
                self.comboBox_ident_instrum_13, self.textEdit_n_serie_instrum_13, self.textEdit_construct_instrum_13, self.textEdit_type_instrum_13, self.textEdit_resolution_instrum_13, self.textEdit_prof_imm_instrum_13, self.radioButton_cofrac_13, self.radioButton_non_cofrac_13, self.textEdit_com_inst_13,
                self.comboBox_ident_instrum_14, self.textEdit_n_serie_instrum_14, self.textEdit_construct_instrum_14, self.textEdit_type_instrum_14, self.textEdit_resolution_instrum_14, self.textEdit_prof_imm_instrum_14, self.radioButton_cofrac_14, self.radioButton_non_cofrac_14, self.textEdit_com_inst_14,
                self.comboBox_ident_instrum_15, self.textEdit_n_serie_instrum_15, self.textEdit_construct_instrum_15, self.textEdit_type_instrum_15, self.textEdit_resolution_instrum_15, self.textEdit_prof_imm_instrum_15, self.radioButton_cofrac_15, self.radioButton_non_cofrac_15, self.textEdit_com_inst_15,
                self.comboBox_ident_instrum_16, self.textEdit_n_serie_instrum_16, self.textEdit_construct_instrum_16, self.textEdit_type_instrum_16, self.textEdit_resolution_instrum_16, self.textEdit_prof_imm_instrum_16, self.radioButton_cofrac_16, self.radioButton_non_cofrac_16, self.textEdit_com_inst_16,
                self.comboBox_ident_instrum_17, self.textEdit_n_serie_instrum_17, self.textEdit_construct_instrum_17, self.textEdit_type_instrum_17, self.textEdit_resolution_instrum_17, self.textEdit_prof_imm_instrum_17, self.radioButton_cofrac_17, self.radioButton_non_cofrac_17, self.textEdit_com_inst_17,
                self.comboBox_ident_instrum_18, self.textEdit_n_serie_instrum_18, self.textEdit_construct_instrum_18, self.textEdit_type_instrum_18, self.textEdit_resolution_instrum_18, self.textEdit_prof_imm_instrum_18, self.radioButton_cofrac_18, self.radioButton_non_cofrac_18, self.textEdit_com_inst_18,
                self.comboBox_ident_instrum_19, self.textEdit_n_serie_instrum_19, self.textEdit_construct_instrum_19, self.textEdit_type_instrum_19, self.textEdit_resolution_instrum_19, self.textEdit_prof_imm_instrum_19, self.radioButton_cofrac_19, self.radioButton_non_cofrac_19, self.textEdit_com_inst_19,
                self.comboBox_ident_instrum_20, self.textEdit_n_serie_instrum_20, self.textEdit_construct_instrum_20, self.textEdit_type_instrum_20, self.textEdit_resolution_instrum_20, self.textEdit_prof_imm_instrum_20, self.radioButton_cofrac_20, self.radioButton_non_cofrac_20, self.textEdit_com_inst_20
                ]
    i=0
    while i < self.SpinBox_nbr_instruments.value():
        
        qlabel[9*i].setCurrentIndex(0)
        
        j =1
        while j <= 5:
            qlabel[9*i +j].clear()
            j+=1
        
        qlabel[9*i + 6].setChecked(False)
        qlabel[9*i + 7].setChecked(False)
        qlabel[9*i + 8].clear()
        
        i+=1
    self.SpinBox_nbr_instruments.setValue(1)
    self.SpinBox_nbr_pt_etal.setValue(1)

def compte_nbr_ligne_textedit_onglet_saisie(self, num_instrum):
    
    list_textedit = [self.textEdit_mesures_inst_1, self.textEdit_mesures_inst_2, self.textEdit_mesures_inst_3, 
                                            self.textEdit_mesures_inst_4, self.textEdit_mesures_inst_5, self.textEdit_mesures_inst_6, 
                                            self.textEdit_mesures_inst_7, self.textEdit_mesures_inst_8, self.textEdit_mesures_inst_9, 
                                            self.textEdit_mesures_inst_10, self.textEdit_mesures_inst_11, self.textEdit_mesures_inst_12, 
                                            self.textEdit_mesures_inst_13, self.textEdit_mesures_inst_14, self.textEdit_mesures_inst_15, 
                                            self.textEdit_mesures_inst_16, self.textEdit_mesures_inst_17, self.textEdit_mesures_inst_18, 
                                            self.textEdit_mesures_inst_19, self.textEdit_mesures_inst_20, 
                                            self.textEdit_mesures_etalon_brutes, self.textEdit_mesures_etalon_corrigees]
                                            
    list_lineedit = [self.nbr_saisie_inst_1, self.nbr_saisie_inst_2, self.nbr_saisie_inst_3, self.nbr_saisie_inst_4, 
                        self.nbr_saisie_inst_5, self.nbr_saisie_inst_6, self.nbr_saisie_inst_7, self.nbr_saisie_inst_8, 
                        self.nbr_saisie_inst_9, self.nbr_saisie_inst_10, self.nbr_saisie_inst_11, self.nbr_saisie_inst_12, 
                        self.nbr_saisie_inst_13, self.nbr_saisie_inst_14, self.nbr_saisie_inst_15, self.nbr_saisie_inst_16,
                        self.nbr_saisie_inst_17, self.nbr_saisie_inst_18, self.nbr_saisie_inst_19, self.nbr_saisie_inst_20, 
                        self.nbr_saisie_etalon_b, self.nbr_saisie_etalon_c]
                        
                        
    a = list_textedit[num_instrum - 1].document()
    list_lineedit[num_instrum - 1].setText("nbr ligne : {}".format(a.lineCount()))
    list_lineedit[num_instrum - 1].setEnabled(False)
