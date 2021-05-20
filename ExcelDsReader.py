# Reading an excel file using Python
from ast import Num
import xlrd
import operator
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import tkinter as tk
from fpdf import FPDF


tkWindow = Tk()  
tkWindow.geometry('400x150')  
tkWindow.title('PythonExamples.org - Tkinter Example')


nameentryframe = Frame(tkWindow, background = 'BLACK', borderwidth = 1)
nameentry = Entry(nameentryframe)
nameentryframe.pack()
nameentry.pack()    
 
  



locExcel="/Users/matthieugrandpierre/Documents/ExcelDSreader/AGIL PRINT SEM21 train de 2.xlsx"
nomDuFichier="test"
CheminPdf="/Users/matthieugrandpierre/Documents"
NumeroTrain=2
"""l"ocExcel=input("Chemin du excel : ")
nomDuFichier=input("Non du pdf : ")
CheminPdf=input("Chemin du pdf : ")
NumeroTrain=int(input("Rentre un numero de train : "))"""



def TrierExcel(locExcel):
    # To open Workbook
    wb = xlrd.open_workbook(locExcel)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)
    DsimpresionData=[]
    
    for i in range(1, sheet.nrows):        
        row = sheet.row_slice(i)        
        TrainId = row[0].value       
        CommunePanneau = row[4].value       
        AdressePanneau = row[5].value   
        Afficheur = row[9].value
        Annonceur1 = row[13].value
        Annonceur2= row[14].value
        Annonceur3= row[15].value
        Annonceur4= row[16].value
        Annonceur5= row[17].value
        Visuel1= row[20].value
        Visuel2= row[23].value
        Visuel3= row[26].value
        Visuel4= row[29].value
        Visuel5= row[32].value

        Produit=[TrainId,CommunePanneau,AdressePanneau,Afficheur,Annonceur1,Visuel1,Annonceur2,Visuel2,Annonceur3,Visuel3,Annonceur4,Visuel4,Annonceur5,Visuel5]
        DsimpresionData.append(Produit) #Tableau de caractéristique des produits

        
    DsimpresionDataTrier = sorted(DsimpresionData, key = operator.itemgetter(5,7,9,11,13))

    return DsimpresionDataTrier

def EcrirePdfTrainDe2(locExcel):
    tmpVisuel1=""
    tmpVisuel2=""
    DsImpressionData=TrierExcel(locExcel)
    monPdf = FPDF()
    monPdf.add_page()
    monPdf.set_font("Arial", size=10)
    compteurPage=0
    compteurPageProduit=0

    compteurProduit=0
    ListecompteurProduit=[]
    i=0
    #tableau des différentes combinaisons de visuel
    for produit in DsImpressionData:
        if (tmpVisuel1!=produit[5] or tmpVisuel2!=produit[7]) and i!=0:
            ListecompteurProduit.append(compteurProduit)  
            compteurProduit=0

        tmpVisuel1=produit[5]
        tmpVisuel2=produit[7]
        compteurProduit+=1
        i+=1
        if i==len(DsImpressionData):
            ListecompteurProduit.append(compteurProduit)  
    j=0

    for ligne in DsImpressionData:

        #si la conbinaison de visuel est différente alors saute de page
        if (tmpVisuel1!=ligne[5] or tmpVisuel2!=ligne[7]) and j!=0:
            monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}-------------------".format(CompteurProduit=ListecompteurProduit[compteurPageProduit],Visuel1=tmpVisuel1,Visuel2=tmpVisuel2), ln=1, align="C")
            compteurPage=0
            compteurPageProduit+=1
        
            monPdf.add_page()

        monPdf.cell(200, 5, txt= "{TrainId}".format(TrainId=ligne[0]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "{CommunePanneau}".format(CommunePanneau=ligne[1]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "{AdressePanneau}".format(AdressePanneau=ligne[2]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "{Afficheur}".format(Afficheur=ligne[3]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "Annonceur1 : {Annonceur1} Visuel1 : {Visuel1}".format(Annonceur1=ligne[4],Visuel1=ligne[5]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "Annonceur2 : {Annonceur2} Visuel2 : {Visuel2}".format(Annonceur2=ligne[6],Visuel2=ligne[7]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "", ln=1, align="L")
        compteurPage+=1
        tmpVisuel1=ligne[5]
        tmpVisuel2=ligne[7]
        j+=1

        #6 produit par page
        if compteurPage==6:
            monPdf.cell(200 ,5, txt="------------------- {CompteurProduit} visuel {Visuel1}/{Visuel2} -------------------".format(CompteurProduit=ListecompteurProduit[compteurPageProduit],Visuel1=ligne[5],Visuel2=ligne[7]), ln=1, align="C")
            compteurPage=0
            monPdf.add_page()
        
        #Le nombre de produit de la conbinaison actuel pour la dernière page
        if j==len(DsImpressionData):
            monPdf.cell(200 ,5, txt="------------------- {CompteurProduit} visuel {Visuel1}/{Visuel2} -------------------".format(CompteurProduit=ListecompteurProduit[compteurPageProduit],Visuel1=ligne[5],Visuel2=ligne[7]), ln=1, align="C")
     
    monPdf.output("{CheminPdf}/{nomDuFichier}.pdf".format(nomDuFichier=nomDuFichier,CheminPdf=CheminPdf))



def EcrirePdfTrainDe3(locExcel):
    tmpVisuel1=""
    tmpVisuel2=""
    tmpVisuel3=""
    DsImpressionData=TrierExcel(locExcel)
    monPdf = FPDF()
    monPdf.add_page()
    monPdf.set_font("Arial", size=10)
    compteurPage=0
    compteurPageProduit=0

    compteurProduit=0
    ListecompteurProduit=[]
    i=0
    for produit in DsImpressionData:
        if (tmpVisuel1!=produit[5] or tmpVisuel2!=produit[7] or tmpVisuel3!=produit[9]) and i!=0:
            ListecompteurProduit.append(compteurProduit)  
            compteurProduit=0

        tmpVisuel1=produit[5]
        tmpVisuel2=produit[7]
        tmpVisuel3=produit[9]
        compteurProduit+=1
        i+=1
        if i==len(DsImpressionData):
            ListecompteurProduit.append(compteurProduit)  
    j=0

    for ligne in DsImpressionData:
        
        if (tmpVisuel1!=ligne[5] or tmpVisuel2!=ligne[7] or tmpVisuel3!=ligne[9]) and j!=0:
            monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}/{Visuel3}-------------------".format(CompteurProduit=ListecompteurProduit[compteurPageProduit],Visuel1=tmpVisuel1,Visuel2=tmpVisuel2,Visuel3=tmpVisuel3), ln=1, align="C")
            compteurPage=0
            compteurPageProduit+=1
            
            monPdf.add_page()

        monPdf.cell(200, 5, txt= "{TrainId}".format(TrainId=ligne[0]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "{CommunePanneau}".format(CommunePanneau=ligne[1]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "{AdressePanneau}".format(AdressePanneau=ligne[2]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "{Afficheur}".format(Afficheur=ligne[3]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "Annonceur1 : {Annonceur1} Visuel1 : {Visuel1}".format(Annonceur1=ligne[4],Visuel1=ligne[5]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "Annonceur2 : {Annonceur2} Visuel2 : {Visuel2}".format(Annonceur2=ligne[6],Visuel2=ligne[7]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "Annonceur3 : {Annonceur3} Visuel3 : {Visuel3}".format(Annonceur3=ligne[8],Visuel3=ligne[9]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "", ln=1, align="L")
        compteurPage+=1
        if compteurPage==6:
            monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}/{Visuel3}-------------------".format(CompteurProduit=ListecompteurProduit[compteurPageProduit],Visuel1=ligne[5],Visuel2=ligne[7],Visuel3=ligne[9]), ln=1, align="C")
            compteurPage=0
            monPdf.add_page()

        tmpVisuel1=ligne[5]
        tmpVisuel2=ligne[7]
        tmpVisuel3=ligne[9]
        j+=1    

        if j==len(DsImpressionData):
            monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}/{Visuel3}-------------------".format(CompteurProduit=ListecompteurProduit[compteurPageProduit],Visuel1=ligne[5],Visuel2=ligne[7],Visuel3=ligne[9]), ln=1, align="C")
     
    monPdf.output("{CheminPdf}/{nomDuFichier}.pdf".format(nomDuFichier=nomDuFichier,CheminPdf=CheminPdf))

def fonction_Principale(locExcel):
    if NumeroTrain==2:
        EcrirePdfTrainDe2(locExcel)
        print("Train de 2")
        
    elif NumeroTrain==3:
        EcrirePdfTrainDe3(locExcel)
        print("Train de 3")

    elif NumeroTrain==4:
        print("Train de 4")

    elif NumeroTrain==5:
        print("Train de 5")
    messagebox.showinfo('Succès', 'Votre fichier a été exporté en pdf avec succès')
    print(nomDuFichier)



def test():
    nomDuFichier=nameentry.get()
    print (nomDuFichier)




button = ttk.Button(tkWindow,
	text = 'Exporter',
	command = lambda: fonction_Principale(locExcel))   
button.pack() 
tkWindow.mainloop()