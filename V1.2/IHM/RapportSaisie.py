
#-*-coding:Latin-1 -*
#import time
import win32com.client
import os
import numpy
#import datetime
import shutil
#----------------------------------------------------------------------
class RapportSaisie():
    '''Class permettant d'exporter les donnees dans la feuille saisie sous excel'''
    
    def __init__(self, nom_campagne):
        self.path = os.path.abspath("AppData/")
        self.feuille_saisie = os.path.abspath("AppData/Documents/rapportsaisie.xlsm")
        self.nom_campagne = nom_campagne
        
        #Afin de pas ecarser le rappor vierge copie du fichier dans le repertoir appdata pour traitement
        shutil.copy2(self.feuille_saisie, self.path)
        
        self.feuille_saisie_travail = os.path.abspath("AppData/rapportsaisie.xlsm")
        
    def mise_en_forme(self, donnees, nom_fichier):
        """"""
        try:
                       
            xl = win32com.client.DispatchEx('Excel.Application')
            classeur = xl.Workbooks.Open(self.feuille_saisie_travail)
            onglet_model = classeur.Worksheets(1)     
            xl.Visible = True
            
            nbr_onglet_total = donnees[0]["nbr_pt_etalonnage"]
            nbr_instrum = len(donnees)
            
            i = nbr_onglet_total
            while i > 0:
    
                nom_onglet = "Pt N�{}__ Temp consigne {}�C".format((i), donnees[0]["temp_consigne"][i-1])
    
                cell = onglet_model.Range("A1:V32")
                cell.Copy()
                
                classeur.Sheets.Add()
    
                #mise en forme de la sheet avec les donnees invariantes pour tous les instrum
                classeur.ActiveSheet.Name = nom_onglet
                classeur.ActiveSheet.Paste()
                classeur.ActiveSheet.Cells(1,1).Value = nom_onglet
                classeur.ActiveSheet.Cells(3,2).Value = donnees[0]["nom_etalon"][i-1]
                
                #Mise en forme ETALON:
                    #mesure_etal_non_corri et moyenne corrigee remarque : on se place sur le premier instrum:
                clef_dico_mesure_etal_non_corri = "mesure_etal_non_corri_pt" + str(i)
                clef_dico_mesure_instrum  = "mesure_instrum_pt"+str(i)
                
                j = 7 #ligne de commencement des donnees dans rapport excel
                for mesure in donnees[0][clef_dico_mesure_etal_non_corri]: #indice 0 car tous les instrume meme mesure d'etalon
                    classeur.ActiveSheet.Cells(j,2).Value = mesure
                    
                    j+=1
                classeur.ActiveSheet.Cells(17,2).Value = donnees[0]["moyenne_etalon_nc"][i-1]
                classeur.ActiveSheet.Cells(18,2).Value = donnees[0]["moyenne_etalon_c"][i-1]
                 
                #Mise en forme GENERATEUR:
                if donnees[0]["nom_generateur"][i-1] == "BGF":
                    #min
                    min = numpy.amin(donnees[0][clef_dico_mesure_etal_non_corri])
                    classeur.ActiveSheet.Cells(30,6).Value = min
                    #max
                    max = numpy.amax(donnees[0][clef_dico_mesure_etal_non_corri])
                    classeur.ActiveSheet.Cells(30,7).Value = max
                    #delta
                    delta = numpy.abs(max - min)
                    classeur.ActiveSheet.Cells(30,8).Value = delta
                    
                elif donnees[0]["nom_generateur"][i-1] == "HART_1":
                    min = numpy.amin(donnees[0][clef_dico_mesure_etal_non_corri])
                    classeur.ActiveSheet.Cells(28,6).Value = min
                    #max
                    max = numpy.amax(donnees[0][clef_dico_mesure_etal_non_corri])
                    classeur.ActiveSheet.Cells(28,7).Value = max
                    #delta
                    delta = numpy.abs(max - min)
                    classeur.ActiveSheet.Cells(28,8).Value = delta
                    
                elif donnees[0]["nom_generateur"][i-1] == "HART_2":
                    min = numpy.amin(donnees[0][clef_dico_mesure_etal_non_corri])
                    classeur.ActiveSheet.Cells(29,6).Value = min
                    #max
                    max = numpy.amax(donnees[0][clef_dico_mesure_etal_non_corri])
                    classeur.ActiveSheet.Cells(29,7).Value = max
                    #delta
                    delta = numpy.abs(max - min)
                    classeur.ActiveSheet.Cells(29,8).Value = delta
                
                elif donnees[0]["nom_generateur"][i-1] == "ESPEC_1":
                    min = numpy.amin(donnees[0][clef_dico_mesure_etal_non_corri])
                    classeur.ActiveSheet.Cells(31,6).Value = min
                    #max
                    max = numpy.amax(donnees[0][clef_dico_mesure_etal_non_corri])
                    classeur.ActiveSheet.Cells(31,7).Value = max
                    #delta
                    delta = numpy.abs(max - min)
                    classeur.ActiveSheet.Cells(31,8).Value = delta
                    
                elif donnees[0]["nom_generateur"][i-1] == "ESPEC_2":
                    min = numpy.amin(donnees[0][clef_dico_mesure_etal_non_corri])
                    classeur.ActiveSheet.Cells(32,6).Value = min
                    #max
                    max = numpy.amax(donnees[0][clef_dico_mesure_etal_non_corri])
                    classeur.ActiveSheet.Cells(32,7).Value = max
                    #delta
                    delta = numpy.abs(max - min)
                    classeur.ActiveSheet.Cells(32,8).Value = delta
                    
                else:
                    min = numpy.amin(donnees[0][clef_dico_mesure_etal_non_corri])
                    classeur.ActiveSheet.Cells(33,6).Value = min
                    #max
                    max = numpy.amax(donnees[0][clef_dico_mesure_etal_non_corri])
                    classeur.ActiveSheet.Cells(33,7).Value = max
                    #delta
                    delta = numpy.abs(max - min)
                    classeur.ActiveSheet.Cells(33,8).Value = delta
                
                #Mise en forme visa
                classeur.ActiveSheet.Cells(27,2).Value = donnees[0]["operateur"]
                
                #Date            
                date = donnees[0]["date_etalonnage"].toString('yyyy-MM-dd') #conversion en string grace � pyqt

                classeur.ActiveSheet.Cells(29,2).Value = date
                
                
                #Campagne
                classeur.ActiveSheet.Cells(31, 2).Value = self.nom_campagne
                
                j=0
                while j < nbr_instrum:
                    
                    classeur.ActiveSheet.Cells(2, (j+3)).Value = donnees[j]["num_document"]
                    classeur.ActiveSheet.Cells(3,(j+3)).Value = donnees[j]["identification_instrument"]
                    
                    resolution = donnees[j]["resolution_instrument"][i-1]
                    classeur.ActiveSheet.Cells(5,(j+3)).Value = resolution
                    
                    classeur.ActiveSheet.Cells(6,(j+3)).Value = donnees[j]["immersion"][i-1]
                    
                    ligne= 7 #debut ligne
                    for ele in donnees[j][clef_dico_mesure_instrum]:
                        classeur.ActiveSheet.Cells(ligne,(j+3)).Value = ele
                        ligne+=1
                        
                    #Moyenne instrum
                    classeur.ActiveSheet.Cells(17,(j+3)).Value = donnees[j]["moyenne_instrum"][i-1]
                    #ecart types
                    classeur.ActiveSheet.Cells(19,(j+3)).Value = donnees[j]["ecart_type"][i-1]
                    #resolution
                    classeur.ActiveSheet.Cells(20,(j+3)).Value = donnees[j]["resolution"][i-1]
                    #stab
                    classeur.ActiveSheet.Cells(21,(j+3)).Value = donnees[j]["stabilite"][i-1]
                    #fuite thermique
                    classeur.ActiveSheet.Cells(22,(j+3)).Value = donnees[j]["fuite_thermique"][i-1]
                    #auto echauffement
                    classeur.ActiveSheet.Cells(23,(j+3)).Value = donnees[j]["autoechauffement"][i-1]
                    #eiv
                    classeur.ActiveSheet.Cells(24,(j+3)).Value = donnees[j]["eiv_meilleure_courante"][i-1]
                    #U
                    classeur.ActiveSheet.Cells(25,(j+3)).Value = donnees[j]["U"][i-1]
                    
                    
                    
                    if nbr_instrum <=10:
                        k = 13
                        while k<= 22:                        
                            selection = classeur.ActiveSheet.Columns(k)
                            selection.Select()
                            selection.Hidden = True
                            k+=1
                    else:
                        k = 13
                        while k<= (22 - nbr_instrum):
                            selection = classeur.ActiveSheet.Columns(k)
                            selection.Select()
                            selection.Hidden = True
                            k+=1                       
    
                
                    j+=1
                
                classeur.ActiveSheet.Protect("TEMP")
                i-=1
                
            xl.CutCopyMode = False
    
            classeur.Sheets("MET-E-020").Visible = False
            
            classeur.SaveAs(self.path + "/"+ nom_fichier)
#            xl.Application.Quit()
            
        except:
            pass
    def fermeture(self):
        
        xl = win32com.client.Dispatch('Excel.Application')
        xl.Application.Quit()
        

