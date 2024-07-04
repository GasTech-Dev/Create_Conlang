import openpyxl
from flask import Flask, request, render_template
import random

class Verbe():
    def trieur(lang):
        #name_ofFile = "Trieur/Franx"#Marquer Le nom de la langue
        name_ofFile = lang
        name_ofFile = name_ofFile + ".xlsx"
        chemin_fichier =name_ofFile
        wb_langue = openpyxl.load_workbook(chemin_fichier)
        #wb_verbe = openpyxl.load_workbook("Trieur/Verbe_Franxois.xlsx")
        wb_verbe = openpyxl.Workbook()
        feuille_langue = wb_langue['Sheet']
        feuille_verbe = wb_verbe['Sheet']
        
        mottr = ""
        i = 1
        
            
        for row in range(1, feuille_langue.max_row + 1):
            MotLangue = feuille_langue.cell(row=row, column=1).value
            MotTraduit = feuille_langue.cell(row=row, column=2).value
            if MotLangue != None:             
                terminaison = MotLangue[-2:]
                if terminaison == "ir" or terminaison == "er":
                    
                    feuille_verbe.cell(row=i, column=2).value = MotLangue
                    feuille_verbe.cell(row=i, column=1).value = MotTraduit
                    print(MotLangue)
                    i += 1
                        

        wb_verbe.save("Verbe_" + lang + ".xlsx")
        wb_langue.close()
        wb_verbe.close()
    def Analyseur(lang):
        #name_ofFile = "Trieur/Verbe_Franxois"#Marquer Le nom de la langue
        print(lang)
        name_ofFile = 'Verbe_' + lang
        name_ofFile = name_ofFile + ".xlsx"
        wb_langue = openpyxl.load_workbook(name_ofFile)
        try:
            feuille_verbe = wb_langue['Sheet']
        except:
            feuille_verbe = wb_langue['Feuil1']
        termot_l = []
        for row in range(1, feuille_verbe.max_row + 1):
            mot = feuille_verbe.cell(row=row, column=1).value
            termot = mot[-2:]
            termot_l.append(termot)

        #print(termot_l)
        first = []
        for i in termot_l:
            if i in first:
                pass
            else:
                first.append(i)
        for i in first:
            f = termot_l.count(i)
            
            if f > 80:
                print(i)
    def Rectifieur(lang):
        #name_ofFile = "Trieur/Verbe_Franxois"#Marquer Le nom de la langue
        name_ofFile = 'Verbe_' + lang
        name_ofFile = name_ofFile + ".xlsx"
        wb_langue = openpyxl.load_workbook(name_ofFile, data_only=True)
        try:
            feuille_verbe = wb_langue['Feuil1']
        except:
            feuille_verbe = wb_langue['Sheet']
        termot_l = []
        for row in range(1, feuille_verbe.max_row + 1):
            mot = feuille_verbe.cell(row=row, column=1).value
            termot = mot[-2:]
            termot_l.append(termot)

        first = []
        for i in termot_l:
            if i in first:
                pass
            else:
                first.append(i)
        terminaison = []
        for i in first:
            f = termot_l.count(i)
            
            if f > 80:
                print(i)
                terminaison.append(i)
        for row in range(1, feuille_verbe.max_row + 1):
            mot = feuille_verbe.cell(row=row, column=1).value
            if mot is not None:
                termot = mot[-2:]
                if termot in terminaison:
                    pass
                else:
                    
                    mot = mot [:-1] + random.choice(terminaison)
                    print(mot)
                    feuille_verbe.cell(row=row, column=1).value = mot
                    #[:-1]
        wb_langue.save("verbe_rectifier_" + lang + ".xlsx")
        wb_langue.close()
class Nom():
    def trieur(lang):
        #name_ofFile = "Trieur/Nom_Franxois"#Marquer Le nom de la langue
        name_ofFile = lang
        name_ofFile = name_ofFile + ".xlsx"
        chemin_fichier =name_ofFile
        wb_langue = openpyxl.load_workbook(chemin_fichier)
        wb_nom = openpyxl.Workbook()
        feuille_langue = wb_langue['Sheet']
        try:
            feuille_nom = wb_nom['Feuil1']
        except:
            feuille_nom = wb_nom['Sheet']
        
        mottr = ""
        i = 1
        
            
        for row in range(1, feuille_langue.max_row + 1):
            MotLangue = feuille_langue.cell(row=row, column=1).value
            MotTraduit = feuille_langue.cell(row=row, column=2).value
            if MotLangue != None:             
                terminaison = MotLangue[-2:]
                if terminaison == "ir" or terminaison=="er":
                    pass
                else:
                    feuille_nom.cell(row=i, column=2).value = MotLangue
                    feuille_nom.cell(row=i, column=1).value = MotTraduit
                    print(MotLangue)
                    i += 1


        wb_nom.save("Nom_" + lang + ".xlsx")
        wb_langue.close()
        wb_nom.close()

    #Annalyse tout les nom et regarde nous donne qu'elle est la terminaison en commun dans chaques nom
    def Analyseur(lang):
        #name_ofFile = "Trieur/Nom_Franxois.xlsx"  # Fusionne les opérations pour éviter une erreur de chemin
        name_ofFile = 'Nom_' + lang + ".xlsx"
        wb_langue = openpyxl.load_workbook(name_ofFile)
        try:
            feuille_verbe = wb_langue['Feuil1']
        except:
            feuille_verbe = wb_langue['Sheet']
        termot_l = []
        
        for row in range(1, feuille_verbe.max_row + 1):
            mot = feuille_verbe.cell(row=row, column=1).value
            
            
            if mot is not None:
                termot = mot[-2:] 
                termot_l.append(termot)
            else:
            
                termot_l.append('')  



        first = []
        for i in termot_l:
            if i in first:
                pass
            else:
                first.append(i)
        for j in first:
            f = termot_l.count(j)
            
            if f > 1200:
                print(j)
                

        wb_langue.save(name_ofFile)
        wb_langue.close()
    def Rectifieur(lang):
        #name_ofFile = "Trieur/Nom_Franxois.xlsx"  # Fusionne les opérations pour éviter une erreur de chemin
        name_ofFile = 'Nom_' + lang + ".xlsx"
        wb_langue = openpyxl.load_workbook(name_ofFile)
        feuille_verbe = wb_langue['Feuil1']
        termot_l = []
        
        for row in range(1, feuille_verbe.max_row + 1):
            mot = feuille_verbe.cell(row=row, column=1).value
            
            
            if mot is not None:
                termot = mot[-2:] 
                termot_l.append(termot)
            else:
            
                termot_l.append('')  



        first = []
        for i in termot_l:
            if i in first:
                pass
            else:
                first.append(i)
        terminaison = []
        for j in first:
            f = termot_l.count(j)
            
            if f > 1200:
                print(j)
                terminaison.append(j)
        for row in range(1, feuille_verbe.max_row + 1):
            mot = feuille_verbe.cell(row=row, column=1).value
            if mot is not None:
                if mot[-2:] in terminaison:
                    pass
                else:
                    mot = mot[:-2] + random.choice(terminaison)
                    feuille_verbe.cell(row=row, column=1).value = mot

        wb_langue.save("Trieur/Nom_Rectifier.xlsx")
        wb_langue.close()
app = Flask(__name__)

@app.route('/trie', methods=['POST'])
def traduct():
    if request.method == 'POST':

        language = request.form.get('language')
        action = request.form.get('action')
        if action == "tv":
            Verbe.trieur(language)
        elif action == "tn":
            Nom.trieur(language)
        elif action == "av":
            Verbe.Analyseur(language)
        elif action == "an":
            Nom.Analyseur(language)
        elif action == "rv":
            Verbe.Rectifieur(language)
        elif action == "rn":
            Nom.Rectifieur(language)

        return render_template('index.html')
    else:

        return render_template('index.html')
@app.route('/')
def index():
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
