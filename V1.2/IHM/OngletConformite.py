#-*-coding:Latin-1 -*
import numpy
from GestionBdd import GestionBdd
import decimal

class OngletConformite():
    '''Classe permettant de gerer l'onglet conformite .
    '''

        
    
    def declaration_conformite(self, n_instrum, nbr_pt_etal, donnees, id_referentiel_conformite, resolution): #attention � voir comment gerer incertitude ou non
        ''' fonction qui permet de verifier la conformit�.
        n_instrum = est un n�
        '''
        
        clef_dico = "moyenne_corrections_instrum_"+str(n_instrum)
        clef_list_resolution = n_instrum -1
        
        resolution_instrument = []
        for ele in resolution[clef_list_resolution].split():
            resolution_instrument.append(float(ele))
        
        resolution_instrum_max = str(numpy.amax(resolution_instrument))    
        nbr_chiffre_significatif = len(resolution_instrum_max[2:]) #on part � 2 pour supprimer la virgule
        
        
        
        #On va chercher le max (|correction| + U):
        correction_incertitude = []
        i=0
        while i < nbr_pt_etal : 
            valeur = round(numpy.abs(donnees[i][clef_dico]), nbr_chiffre_significatif) + float(decimal.Decimal(str(donnees[i]["U"])).quantize(decimal.Decimal(resolution_instrum_max), rounding=decimal.ROUND_UP) )   #round(float(donnees[i]["U"]), nbr_chiffre_significatif)
            correction_incertitude.append(valeur)
            
            i+=1
        max_correction_incertitude = round(numpy.amax(correction_incertitude), nbr_chiffre_significatif)
        
        #on recupere dans la base la valeur de l'EMT 
        #Acces bdd
        db = GestionBdd('db')
        db.reconnexion()
        valeur_emt = db.recuperation_valeur_emt(id_referentiel_conformite)
#        print("valeur EMT {}".format(valeur_emt))

        if max_correction_incertitude <= valeur_emt :
            resultat = "Conforme"
            valeur = "{} <= {}".format(max_correction_incertitude, valeur_emt)
            
        else:
            resultat = "Non Conforme"
            valeur = "{} >= {}".format(max_correction_incertitude, valeur_emt)
            
        return resultat, valeur
