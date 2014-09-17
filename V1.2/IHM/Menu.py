# -*- coding: utf-8 -*-

"""
Module implementing MainWindow.
"""

from PyQt4.QtCore import pyqtSlot
from PyQt4.QtGui import QMainWindow
from PyQt4 import QtGui, QtCore
from PyQt4.QtGui import QFileDialog
from PyQt4.QtGui import QMessageBox
from PyQt4.QtGui import QInputDialog
from Ui_Menu import Ui_Menu
from Saisie_etalonnage import Saisie
from GestionBdd import GestionBdd
from RapportEtalonnage import RapportEtalonnage
from RapportSaisie import RapportSaisie
from Package.fonctions import sauvegarde_onglet_saisie
import numpy
import os



class Menu(QMainWindow, Ui_Menu):
    """
    Class documentation goes here.
    """
    def __init__(self, parent=None):
        """ Constructor @param parent reference to the parent widget (QWidget)
        """
        super().__init__(parent)
        self.setupUi(self)
         # acces à la BDD
        db = GestionBdd('db')
#        db.premiere_connexion()
        db.reconnexion()
        
    @pyqtSlot()
    def on_buttonBox_accepted(self):
        """ Slot documentation goes here.
        """
        # TODO: not implemented yet
        raise NotImplementedError
    
    @pyqtSlot()
    def on_radioButton_saisie_etal_clicked(self):
        """
        ouvre le qmainwindows saisie
        """     
        reponse = QMessageBox.question(self, 
                self.trUtf8("Information"), 
                self.trUtf8("Voulez-vous saisir un nouvel étalonnage"), 
                QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)        
               
        if reponse == QtGui.QMessageBox.Yes:            
            self.saisie = Saisie([1])
            self.saisie.setWindowModality(QtCore.Qt.ApplicationModal)
            self.saisie.show()
        else :
            pass
        
    @pyqtSlot()
    def on_radioButton_reeimpression_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #Acces bdd
        db = GestionBdd('db')
        db.reconnexion()
        
        valeur_combobox = db.recuperation_n_ce()
        self.comboBox_reeimp_n_document.addItems(valeur_combobox["num_doc"])
        
        
    
    @pyqtSlot()
    def on_radioButton_annule_remplace_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        #Acces bdd
        db = GestionBdd('db')
        db.reconnexion()
        
        valeur_combobox = db.recuperation_n_ce()
        self.comboBox_annule_remplace_n_doc.addItems(valeur_combobox["num_doc"])
    
    @pyqtSlot()
    def on_radioButton_validation_clicked(self):
        """
        Slot documentation goes here.
        """
        reponse = QMessageBox.question(self, 
        self.trUtf8("Information"), 
        self.trUtf8("Voulez-vous rechercher par n° campagne d'étalonnage ? \n"\
                        +"Cliquer non pour rechercher par n° certificat"), 
        QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)        
               
        if reponse == QtGui.QMessageBox.Yes:  
            db = GestionBdd('db')
            db.reconnexion()            
            ensemble_nom_campagne = db.recuperation_nom_campagne_etal_non_valide()
            
            valeur_combobox = []
            for ele in ensemble_nom_campagne:
                valeur_combobox.append("Campagne d'étalonnage n° {}".format(ele))
            self.comboBox_validation_n_doc.addItems(valeur_combobox)
            
        else:
            db = GestionBdd('db')
            db.reconnexion()
            ensemble_nom_etalonnage = db.recuperation_n_ce_non_valide() #attention il s'agit d'un dico avec id_campagne_etal ,num_doc comme clef
            valeur_combobox = []
            for ele in ensemble_nom_etalonnage["num_doc"]:
                valeur_combobox.append("Etalonnage n° {}".format(ele))
#            print(valeur_combobox)
            self.comboBox_validation_n_doc.addItems(valeur_combobox)
            
    
    @pyqtSlot(str)
    def on_comboBox_reeimp_n_document_activated(self, p0):
        """
       Fct qui recupere l'ensemble des donnees d'une etalonnage
        """
        # TODO: not implemented yet
        #Acces bdd
        db = GestionBdd('db')
        db.reconnexion()        
        n_certificat = self.comboBox_reeimp_n_document.currentText()
                
        ce = db.recuperation_donnees_etalonnage(n_certificat)
        cv = db.recuperation_donnees_conformite(ce)                
        rapport_etalonnage = dict(list(ce.items()) + list(cv.items()))
        
        list_etalon = list(set(rapport_etalonnage["etalon"]))
        list_generateur = list(set(rapport_etalonnage["generateur"]))
        rapport_etalonnage["etalon"] = list_etalon
        rapport_etalonnage["generateur"] = list_generateur
        
        
        #on choisit le dossier de sauvegarde
        dossier = QFileDialog.getExistingDirectory(None ,  "Selectionner le dossier de sauvegarde des Rapports")
        
        #gestion n_ mode operatoire
        if rapport_etalonnage["milieu"] == "LIQUIDE":
            rapport_etalonnage["n_mode_operatoire"] = "025"
        else :
            rapport_etalonnage["n_mode_operatoire"] = "030"
        
        date_etal = rapport_etalonnage["date_etalonnage"].toString('dd/MM/yyyy')
        rapport_etalonnage["date_etalonnage"] = date_etal
        
        if isinstance(rapport_etalonnage["renseignement_complementaire"], str):
            pass
        else:
            rapport_etalonnage["renseignement_complementaire"] = ""
            
        if not rapport_etalonnage["Annule_doc"]:
            certificat = RapportEtalonnage(rapport_etalonnage["CE"]) #appel à la classe RapportEtalonnage
            certificat.mise_en_forme_ce(rapport_etalonnage, dossier, rapport_etalonnage["n_certificat"])
        else:
            nom_fichier_ce =   rapport_etalonnage["n_certificat"] + " Annule et Remplace le document "+ rapport_etalonnage["Annule_doc"]
            nom_ce = rapport_etalonnage["n_certificat"] + "\n Annule et Remplace le document "+ rapport_etalonnage["Annule_doc"]
            certificat = RapportEtalonnage(rapport_etalonnage["CE"]) #appel à la classe RapportEtalonnage
            certificat.mise_en_forme_ce_annule_remplace(rapport_etalonnage, dossier, nom_ce, nom_fichier_ce)
            
            
        if not cv["conformite"]:
            pass            
        else:
            if not rapport_etalonnage["Annule_doc"]:            
                nom_fichier_constat = str(rapport_etalonnage["n_certificat"]) + "V"
                constat = RapportEtalonnage( rapport_etalonnage["CE"])
                constat.mise_en_forme_cv(rapport_etalonnage, dossier,  nom_fichier_constat)
            else:
                nom_cv =   rapport_etalonnage["n_certificat"] + "V" + "\n Annule et Remplace le document "+ rapport_etalonnage["Annule_doc"] +"V"
                nom_fichier_cv = rapport_etalonnage["n_certificat"] + "V" + " Annule et Remplace le document "+ rapport_etalonnage["Annule_doc"] +"V"
                
                constat = RapportEtalonnage(rapport_etalonnage["CE"])
                constat.mise_en_forme_cv_annule_remplace(rapport_etalonnage, dossier, nom_cv, nom_fichier_cv)
                
        self.comboBox_reeimp_n_document.clear()
        
    @pyqtSlot(str)
    def on_comboBox_annule_remplace_n_doc_activated(self, p0):
        """
        Slot documentation goes here.
        """
      
        n_document_selectionne = self.comboBox_annule_remplace_n_doc.currentText()
        
        db = GestionBdd('db')
        db.reconnexion()           
        id_campagne_etal = db.recherche_id_campagne_etal_n_etalonnage(n_document_selectionne)
        donnees_campagne = db.recuperation_donnees_etalonnage_validation(id_campagne_etal) #on recupere toute les donnees de la campagne
        
        donnees_n_document_selectionne = []
        for ele in donnees_campagne: # on trie les données de la campagne pour ne garder que le n°certificat selectionné
            if ele["num_document"] == n_document_selectionne:
                donnees_n_document_selectionne.append(ele)
                   
#
#        rapport = RapportSaisie()
#        rapport.mise_en_forme(donnees_n_document_selectionne, n_document_selectionne)
#
#        print("donnees passées par le menu annule et remplace {}".format(donnees_n_document_selectionne))              
        self.gestion_fichier_pickler_saisie_etalonnage(donnees_n_document_selectionne)
        self.saisie = Saisie([4, n_document_selectionne, id_campagne_etal])
        self.saisie.setWindowModality(QtCore.Qt.ApplicationModal)
        self.saisie.show()
    
    
    @pyqtSlot(str)
    def on_comboBox_validation_n_doc_activated(self, p0):
        """
        Appel la fct recuperation des donnees d etal de la classe bdd
        et export ces donnees dans un classeur excel class Rapport Saisie
        """
        
        valeur_selectionnee_combobox = self.comboBox_validation_n_doc.currentText()
        
        if valeur_selectionnee_combobox[0] == "C": #test pour savoir s'il s'agit d'un n° certificat ou de caompagne d'etalonnage            
           nom_campagne_etal_particuliere = valeur_selectionnee_combobox[25:len(valeur_selectionnee_combobox)]
           
           db = GestionBdd('db')
           db.reconnexion()
            
           id_campagne_etal = db.recherche_id_campagne_etal_particuliere(nom_campagne_etal_particuliere)
            
           donnees_campagne = db.recuperation_donnees_etalonnage_validation(id_campagne_etal)
#           print(donnees_campagne)
           rapport = RapportSaisie(nom_campagne_etal_particuliere)
           rapport.mise_en_forme(donnees_campagne, nom_campagne_etal_particuliere)
            
           reponse = QMessageBox.question(self, 
                    self.trUtf8("Information"), 
                    self.trUtf8("Voulez-vous archiver cette campagne d'étalonnage"), 
                    QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
            
                   
           if reponse == QtGui.QMessageBox.Yes:
                db.archivage_campage_etal(id_campagne_etal)
                db.archivage_etalonnage(id_campagne_etal)
                ensemble_nom_campagne = db.recuperation_nom_campagne_etal_non_valide()
                valeur_combobox = [nom for nom in ensemble_nom_campagne if bool(nom)!= False]
                self.comboBox_validation_n_doc.clear()
                #effacement des fichiers temporaire (present dans le dossier AppData) 
            
                path =os.path.abspath("AppData/")
          
                for ele in os.listdir(path):
                    path_total = str(path + "/"+str(ele))
                
                    if os.path.isfile(path_total): # verification qu'il s'agit bien de fichier
                        os.remove(path_total)
#                self.comboBox_validation_n_doc.addItems(valeur_combobox)
                
           else:
                reponse = QMessageBox.question(self, 
                    self.trUtf8("Information"), 
                    self.trUtf8("Voulez-vous modifier les donnees?"), 
                    QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
                    
                if reponse == QtGui.QMessageBox.Yes:
                    rapport.fermeture()
                    instrument = []
#                    QMessageBox.information(self, 
#                    self.trUtf8("Information"), 
#                    self.trUtf8("Merci de fermer le rapport de saisie (fichier excel)"))                     

                    self.gestion_fichier_pickler_saisie_etalonnage(donnees_campagne)
                    self.saisie = Saisie([2, nom_campagne_etal_particuliere, id_campagne_etal])
                    self.saisie.setWindowModality(QtCore.Qt.ApplicationModal)
                    self.saisie.show()
        else:
            n_document_selectionne = valeur_selectionnee_combobox[14:len(valeur_selectionnee_combobox)]
        
            db = GestionBdd('db')
            db.reconnexion()           
            id_campagne_etal = db.recherche_id_campagne_etal_n_etalonnage(n_document_selectionne)
            donnees_campagne = db.recuperation_donnees_etalonnage_validation(id_campagne_etal) #on recupere toute les donnees de la campagne
            nom_campagne_etal_particuliere = db.recherche_nom_campagne_etal_id_campagne(id_campagne_etal)
            
            
            id_etalonnage = db.recherche_id_etalonnage_num_doc(n_document_selectionne)
            
            
            
            donnees_n_document_selectionne = []
            for ele in donnees_campagne: # on trie les données de la campagne pour ne garder que le n°certificat selectionné
                if ele["num_document"] == n_document_selectionne:
                    donnees_n_document_selectionne.append(ele)
                       

            rapport = RapportSaisie(nom_campagne_etal_particuliere)
            rapport.mise_en_forme(donnees_n_document_selectionne, n_document_selectionne)
            
            reponse = QMessageBox.question(self, 
                    self.trUtf8("Information"), 
                    self.trUtf8("Voulez-vous archiver cet d'étalonnage"), 
                    QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
            
                   
            if reponse == QtGui.QMessageBox.Yes:
                db.archivage_etalonnage_unique(id_etalonnage)
                ensemble_nom_campagne = db.recuperation_nom_campagne_etal_non_valide()
                self.comboBox_validation_n_doc.clear()
#                self.comboBox_validation_n_doc.addItems(valeur_combobox)
                
            else:
                reponse = QMessageBox.question(self, 
                    self.trUtf8("Information"), 
                    self.trUtf8("Voulez-vous modifier les donnees?"), 
                    QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
                if reponse == QtGui.QMessageBox.Yes:                                
                    rapport.fermeture()
#                    QMessageBox.information(self, 
#                    self.trUtf8("Information"), 
#                    self.trUtf8("Merci de fermer le rapport de saisie (fichier excel)")) 
                    

                    self.gestion_fichier_pickler_saisie_etalonnage(donnees_n_document_selectionne)
                    self.saisie = Saisie([3, n_document_selectionne, nom_campagne_etal_particuliere])
                    self.saisie.setWindowModality(QtCore.Qt.ApplicationModal)
                    self.saisie.show()

            
                
    def gestion_fichier_pickler_saisie_etalonnage(self, donnees_campagne):
        ''''fonction qui gere la creation des fichiers pickler en vue de les modifier
        dans Saisie_Etalonnage'''
        
        #config:
        db = GestionBdd('db')
        db.reconnexion()
        fichier_configuration =[]
        config_etal = {}
        fichier_saisie = {}
               
        nbr_instrum = len(donnees_campagne)
        config_etal["nbr_instrum"] = nbr_instrum
        
        nbr_pt_etal = donnees_campagne[0]["nbr_pt_etalonnage"]
        config_etal["nbr_pts_temp"] = nbr_pt_etal
        
        date_1 = donnees_campagne[0]["date_etalonnage"]
       
#        date = QtCore.QDate.fromString(date_1, "yyyy-MM-dd")
        
        config_etal["Date"] = date_1
        
        fichier_configuration.append(config_etal)
        
        i=0
        while i < nbr_instrum:
            instrument = {}
            resolution = []
            
            instrum = db.gestion_remplissage_onglet_config_2(donnees_campagne[i]["identification_instrument"])
            instrument["nom_instrum"] = donnees_campagne[i]["identification_instrument"]
            
            instrument["n_serie"] = instrum[0]
            instrument["constructeur"] = instrum[1]
            instrument["Type"] = instrum[2]
            
#            resolution.append(instrum[3])
            instrument["resolution"] = donnees_campagne[i]["resolution_instrument"]
            
            if isinstance(instrum[4], str):
                instrument["renseignement_complementaire"] = instrum[4]
            else:
                instrument["renseignement_complementaire"] = ""
            
            if donnees_campagne[i]["CE"] == "COFRAC":
                instrument["Type_etalonnage"] ="Cofrac" 
            else:
                instrument["Type_etalonnage"] ="Non Cofrac" 
                
            instrument["Etat_reception"] = donnees_campagne[i]["Etat_reception"]
            instrument["immersion"] = donnees_campagne[i]["immersion"]                
            fichier_configuration.append(instrument)
            
            
            
            i+=1
            
        sauvegarde_onglet_saisie(fichier_configuration , "configuration")
       
        #gestion fichiers saisies
                #pt 1
        i=1
        while i <= nbr_pt_etal:
            j=0
            while j < nbr_instrum:
                                   
                fichier_saisie["mesures_inst_"+str(j+1)] = donnees_campagne[j]["mesure_instrum_pt"+str(i)]
                fichier_saisie["mesures_etal_brute"] = donnees_campagne[j]["mesure_etal_non_corri_pt"+str(i)]
                fichier_saisie["mesures_etal_corri"] = donnees_campagne[j]["mesure_etal_corri_pt"+str(i)]
                fichier_saisie["corrections_instrum_"+str(j+1)] = numpy.array(donnees_campagne[j]["mesure_instrum_pt"+str(i)])
                fichier_saisie["nb_acquisition"] = len(donnees_campagne[j]["mesure_instrum_pt"+str(i)])
                
                #recherche id operteur
                operateur = [] 
                
                for ele in donnees_campagne[j]["operateur"].split() :                
                    operateur.append(ele)
                prenom = operateur[0]
                nom =  operateur[1]
                id_operateur = db.recuperation_id_operateur(prenom, nom)
                
                fichier_saisie["operateur"] = id_operateur -1 #combobox partent à 0
                
                
                fichier_saisie["auto_echauffement_instrum_"+str(j)] = donnees_campagne[j]["autoechauffement"][i-1]
                
                if donnees_campagne[j]["nom_etalon"][i-1] == "THE REF 003":
                    fichier_saisie["etalon"] = 0 #attention il nous faut un numero
                elif donnees_campagne[j]["nom_etalon"][i-1] == "THE REF 004":
                    fichier_saisie["etalon"] = 1
                elif donnees_campagne[j]["nom_etalon"][i-1] == "THE REF 005": 
                    fichier_saisie["etalon"] = 2                    
                elif donnees_campagne[j]["nom_etalon"][i-1] == "THE REF 006":
                    fichier_saisie["etalon"] = 3                    
                elif donnees_campagne[j]["nom_etalon"][i-1] == "THE REF 007":
                    fichier_saisie["etalon"] = 4                        
                elif donnees_campagne[j]["nom_etalon"][i-1] == "THE REF 008":
                    fichier_saisie["etalon"] = 5
                    
                if donnees_campagne[j]["nom_generateur"][i-1] == "BGF":
                    fichier_saisie["generateur"] = 0 #attention il nous faut un numero
                elif donnees_campagne[j]["nom_generateur"][i-1] == "HART_1":
                    fichier_saisie["generateur"] = 1
                elif donnees_campagne[j]["nom_generateur"][i-1] == "HART_2": 
                    fichier_saisie["generateur"] = 2                    
                elif donnees_campagne[j]["nom_generateur"][i-1] == "ESPEC_1":
                    fichier_saisie["generateur"] = 3                    
                elif donnees_campagne[j]["nom_generateur"][i-1] == "ESPEC_2":
                    fichier_saisie["generateur"] = 4                        
                
               
                fichier_saisie["moyenne_instrum_"+str(j+1)] = donnees_campagne[j]["moyenne_instrum"][i-1]
                fichier_saisie["ecartype_corrections_instrum_"+str(j+1)] = donnees_campagne[j]["ecart_type"][i-1]
                fichier_saisie["temp_consig"] = donnees_campagne[j]["temp_consigne"][i-1]
                fichier_saisie["moyenne_corrections_instrum_"+str(j)] = donnees_campagne[j]["moyenne_correction"][i-1]
                fichier_saisie["moyenne_etal_corri"] = donnees_campagne[j]["moyenne_etalon_c"][i-1]
                fichier_saisie["fuite_thermique_instrum_"+str(j+1)] = donnees_campagne[j]["fuite_thermique"][i-1]
                
                j+=1
                
            fichier_saisie["pt_etalonnag_n"] = i
            fichier_saisie["chemin_fichier_etalon"] = "Neant"
            sauvegarde_onglet_saisie(fichier_saisie,i )
            i+=1
    
    @pyqtSlot()
    def on_radioButton_consultation_clicked(self):
        """
        Slot documentation goes here.
        """
        reponse = QMessageBox.question(self, 
        self.trUtf8("Information"), 
        self.trUtf8("Voulez-vous rechercher par n° campagne d'étalonnage ? \n"\
                        +"Cliquer non pour rechercher par n° certificat"), 
        QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)        
               
        if reponse == QtGui.QMessageBox.Yes:  
            db = GestionBdd('db')
            db.reconnexion()            
            ensemble_nom_campagne = db.recuperation_nom_campagne_etal_all()
            
            valeur_combobox = []
            for ele in ensemble_nom_campagne:
                valeur_combobox.append("Campagne d'étalonnage n° {}".format(ele))
            self.comboBox_consultation.addItems(valeur_combobox)
            
        else:
            db = GestionBdd('db')
            db.reconnexion()
            ensemble_nom_etalonnage = db.recuperation_n_ce_all() #attention il s'agit d'un dico avec id_campagne_etal ,num_doc comme clef
            valeur_combobox = []
            for ele in ensemble_nom_etalonnage["num_doc"]:
                valeur_combobox.append("Etalonnage n° {}".format(ele))
#            print(valeur_combobox)
            self.comboBox_consultation.addItems(valeur_combobox)
    
    @pyqtSlot(str)
    def on_comboBox_consultation_activated(self, p0):
        """
        Slot documentation goes here.
        """
      
        valeur_selectionnee_combobox = self.comboBox_consultation.currentText()
        
        if valeur_selectionnee_combobox[0] == "C": #test pour savoir s'il s'agit d'un n° certificat ou de caompagne d'etalonnage            
           nom_campagne_etal_particuliere = valeur_selectionnee_combobox[25:len(valeur_selectionnee_combobox)]
           
           db = GestionBdd('db')
           db.reconnexion()
            
           id_campagne_etal = db.recherche_id_campagne_etal_particuliere(nom_campagne_etal_particuliere)
            
           donnees_campagne = db.recuperation_donnees_etalonnage_validation(id_campagne_etal)
    #        print(donnees_campagne)
           rapport = RapportSaisie(nom_campagne_etal_particuliere)
           rapport.mise_en_forme(donnees_campagne, nom_campagne_etal_particuliere)
           
           
        else:
            n_document_selectionne = valeur_selectionnee_combobox[14:len(valeur_selectionnee_combobox)]
        
            db = GestionBdd('db')
            db.reconnexion()           
            id_campagne_etal = db.recherche_id_campagne_etal_n_etalonnage(n_document_selectionne)
            donnees_campagne = db.recuperation_donnees_etalonnage_validation(id_campagne_etal) #on recupere toute les donnees de la campagne
            nom_campagne = db.recherche_nom_campagne_etal_id_campagne(id_campagne_etal)
            donnees_n_document_selectionne = []
            for ele in donnees_campagne: # on trie les données de la campagne pour ne garder que le n°certificat selectionné
                if ele["num_document"] == n_document_selectionne:
                    donnees_n_document_selectionne.append(ele)
                       

            rapport = RapportSaisie(nom_campagne)
            rapport.mise_en_forme(donnees_n_document_selectionne, n_document_selectionne)
