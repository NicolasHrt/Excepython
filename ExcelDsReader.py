# Reading an excel file using Python
from ast import Num
from os import name
import xlrd
import operator
from fpdf import FPDF
from tkinter import ttk
from tkinter import *
from tkinter import messagebox
# import filedialog module
from tkinter import filedialog

# Create the root window
window = Tk()

# Set window title
window.title("ExcelDsPrintConverter")

# Set window size
window.geometry("1100x1000")

NumeroTrain=IntVar()
locExcel=""
CheminPdf=""

# Function for opening thes
# file explorer window
def BrowseExcel():
    global locExcel
    locExcel = filedialog.askopenfilename(initialdir = "/",
                                            title = "Select a File",
                                            filetypes=[("Excel files", ".xlsx .xls")])
        
    label_file_explorer.configure(text="File Opened: "+locExcel)

def OutputPdf():
    global CheminPdf
    CheminPdf = filedialog.askdirectory(initialdir = "/",
										title = "Select a File",)

    label_file_output.configure(text="File Opened: "+CheminPdf)

def TrierExcel(locExcel):
    # To open Workbook
    wb = xlrd.open_workbook(locExcel)
    sheet = wb.sheet_by_index(int(NumeroFeuille.get())-1)
    print(sheet.name)
    sheet.cell_value(0, 0)
    DsimpresionData=[]
    errorChampVide()    
    for i in range(1, sheet.nrows):   
        row = sheet.row_slice(i)        
        TrainId = row[int(ColonneTrainId.get())-1].value  
        #TrainId = row[0].value     
        CommunePanneau = row[int(ColonneCommune.get())-1].value       
        AdressePanneau = row[int(ColonneAdresse.get())-1].value   
        Afficheur = row[int(ColonneAfficheur.get())-1].value
        Annonceur1 = row[int(ColonneAnnonceur1.get())-1].value
        Visuel1= row[int(ColonneVisuel1.get())-1].value
       

        if NumeroTrain.get()==1:
            Produit=[TrainId,CommunePanneau,AdressePanneau,Afficheur,Annonceur1,Visuel1]

        if NumeroTrain.get()==2:
            Annonceur2= row[int(ColonneAnnonceur2.get())-1].value
            Visuel2= row[int(ColonneVisuel2.get())-1].value
            Produit=[TrainId,CommunePanneau,AdressePanneau,Afficheur,Annonceur1,Visuel1,Annonceur2,Visuel2]


        if NumeroTrain.get()==3:
            Annonceur2= row[int(ColonneAnnonceur2.get())-1].value
            Annonceur3= row[int(ColonneAnnonceur3.get())-1].value
            Visuel2= row[int(ColonneVisuel2.get())-1].value
            Visuel3= row[int(ColonneVisuel3.get())-1].value
            Produit=[TrainId,CommunePanneau,AdressePanneau,Afficheur,Annonceur1,Visuel1,Annonceur2,Visuel2,Annonceur3,Visuel3]

        if NumeroTrain.get()==4:
            Annonceur2= row[int(ColonneAnnonceur2.get())-1].value
            Annonceur3= row[int(ColonneAnnonceur3.get())-1].value
            Annonceur4= row[int(ColonneAnnonceur4.get())-1].value
            Visuel2= row[int(ColonneVisuel2.get())-1].value
            Visuel3= row[int(ColonneVisuel3.get())-1].value
            Visuel4= row[int(ColonneVisuel4.get())-1].value
            Produit=[TrainId,CommunePanneau,AdressePanneau,Afficheur,Annonceur1,Visuel1,Annonceur2,Visuel2,Annonceur3,Visuel3,Annonceur4,Visuel4]

        if NumeroTrain.get()==5:
            Annonceur2= row[int(ColonneAnnonceur2.get())-1].value
            Annonceur3= row[int(ColonneAnnonceur3.get())-1].value
            Annonceur4= row[int(ColonneAnnonceur4.get())-1].value
            Annonceur5= row[int(ColonneAnnonceur5.get())-1].value
            Visuel2= row[int(ColonneVisuel2.get())-1].value
            Visuel3= row[int(ColonneVisuel3.get())-1].value
            Visuel4= row[int(ColonneVisuel4.get())-1].value
            Visuel5= row[int(ColonneVisuel5.get())-1].value
            Produit=[TrainId,CommunePanneau,AdressePanneau,Afficheur,Annonceur1,Visuel1,Annonceur2,Visuel2,Annonceur3,Visuel3,Annonceur4,Visuel4,Annonceur5,Visuel5]
    
        DsimpresionData.append(Produit) #Tableau de caractéristique des produits

    DsimpresionDataTrier=fonctionDeTrie(DsimpresionData)


    #print(DsimpresionDataTrier)
    return DsimpresionDataTrier

def errorChampVide():
    if ColonneTrainId.get()=="":messagebox.showerror("Error", "Vous n'avez pas rempli le champ trainId")
    elif ColonneCommune.get()=="":messagebox.showerror("Error", "Vous n'avez pas rempli le champ Commune du panneau")
    elif ColonneAdresse.get()=="":messagebox.showerror("Error", "Vous n'avez pas rempli le champ Adresse du panneau")
    elif ColonneAfficheur.get()=="":messagebox.showerror("Error", "Vous n'avez pas rempli le champ Afficheur")
    elif ColonneAnnonceur1.get()=="":messagebox.showerror("Error", "Vous n'avez pas rempli le champ Annonceur 1")
    elif ColonneVisuel1.get()=="":messagebox.showerror("Error", "Vous n'avez pas rempli le champ Visuel 1")
    elif (NumeroTrain.get()==2 or NumeroTrain.get()==3 or NumeroTrain.get()==4 or NumeroTrain.get()==5) and ColonneAnnonceur2.get()=="":messagebox.showerror("Error", "Vous n'avez pas rempli le champ Annonceur 2")
    elif (NumeroTrain.get()==2 or NumeroTrain.get()==3 or NumeroTrain.get()==4 or NumeroTrain.get()==5) and ColonneVisuel2.get()=="":messagebox.showerror("Error", "Vous n'avez pas rempli le champ Visuel 2")
    elif (NumeroTrain.get()==3 or NumeroTrain.get()==4 or NumeroTrain.get()==5)  and ColonneAnnonceur3.get()=="":messagebox.showerror("Error", "Vous n'avez pas rempli le champ Annonceur 3")
    elif (NumeroTrain.get()==3 or NumeroTrain.get()==4 or NumeroTrain.get()==5) and ColonneVisuel3.get()=="":messagebox.showerror("Error", "Vous n'avez pas rempli le champ Visuel 3")
    elif (NumeroTrain.get()==4 or NumeroTrain.get()==5) and ColonneAnnonceur4.get()=="":messagebox.showerror("Error", "Vous n'avez pas rempli le champ Annonceur 4")
    elif (NumeroTrain.get()==4 or NumeroTrain.get()==5 )and ColonneVisuel4.get()=="":messagebox.showerror("Error", "Vous n'avez pas rempli le champ Visuel 4")
    elif NumeroTrain.get()==5 and ColonneAnnonceur5.get()=="":messagebox.showerror("Error", "Vous n'avez pas rempli le champ Annonceur 5")
    elif NumeroTrain.get()==5 and ColonneVisuel5.get()=="":messagebox.showerror("Error", "Vous n'avez pas rempli le champ Visuel 5")
    else:pass

def fonctionDeTrie(DsimpresionData):
    if NumeroTrain.get()==1:  
        DsimpresionDataTrier = sorted(DsimpresionData, key = operator.itemgetter(5))

    if NumeroTrain.get()==2:  
        DsimpresionDataTrier = sorted(DsimpresionData, key = operator.itemgetter(5,7))
    
    if NumeroTrain.get()==3:  
        DsimpresionDataTrier = sorted(DsimpresionData, key = operator.itemgetter(5,7,9))
    
    if NumeroTrain.get()==4:  
        DsimpresionDataTrier = sorted(DsimpresionData, key = operator.itemgetter(5,7,9,11))
    
    if NumeroTrain.get()==5:  
        DsimpresionDataTrier = sorted(DsimpresionData, key = operator.itemgetter(5,7,9,11,13))
    return(DsimpresionDataTrier)

def EcrirePdfTrainDe1(locExcel):
    tmpVisuel1=""
    DsImpressionData=TrierExcel(locExcel)
    monPdf = FPDF()
    monPdf.add_page()
    monPdf.set_font("Arial", size=20)
    monPdf.cell(200 ,5, txt="Liste recapitulative de toutes les combinaisons de visuels", ln=1, align="C") 
    monPdf.cell(200 ,10, txt="", ln=2, align="C") 
    monPdf.set_font("Arial", size=10) 
    compteurPage=0
    compteurPageProduit=0

    compteurProduit=0
    ListecompteurProduit=[]
    i=0

    #tableau des différentes combinaisons de visuel
    for produit in DsImpressionData:
        if tmpVisuel1!=produit[5] and i!=0:
            ListecompteurProduit.append(compteurProduit)
            monPdf.cell(200 ,5, txt="------------------- {CompteurProduit} visuel {Visuel1} -------------------".format(CompteurProduit=compteurProduit,Visuel1=tmpVisuel1), ln=1, align="C")  
            compteurProduit=0

        tmpVisuel1=produit[5]
        compteurProduit+=1
        i+=1
        if i==len(DsImpressionData):
            ListecompteurProduit.append(compteurProduit)  
            monPdf.cell(200 ,5, txt="------------------- {CompteurProduit} visuel {Visuel1} -------------------".format(CompteurProduit=compteurProduit,Visuel1=produit[5]), ln=1, align="C")  
            monPdf.add_page()
    j=0

    for ligne in DsImpressionData:

        #si la conbinaison de visuel est différente alors saute de page
        if tmpVisuel1!=ligne[5] and j!=0:
            monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}-------------------".format(CompteurProduit=ListecompteurProduit[compteurPageProduit],Visuel1=tmpVisuel1), ln=1, align="C")
            compteurPage=0
            compteurPageProduit+=1
            monPdf.add_page()
        
        monPdf.cell(200, 5, txt= "{TrainId}".format(TrainId=ligne[0]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "{CommunePanneau}".format(CommunePanneau=ligne[1]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "{AdressePanneau}".format(AdressePanneau=ligne[2]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "{Afficheur}".format(Afficheur=ligne[3]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "Annonceur1 : {Annonceur1} Visuel1 : {Visuel1}".format(Annonceur1=ligne[4],Visuel1=ligne[5]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "", ln=1, align="L")
        compteurPage+=1
        tmpVisuel1=ligne[5]
        j+=1

        #6 produit par page
        if compteurPage==6:
            monPdf.cell(200 ,5, txt="------------------- {CompteurProduit} visuel {Visuel1} -------------------".format(CompteurProduit=ListecompteurProduit[compteurPageProduit],Visuel1=ligne[5]), ln=1, align="C")
            compteurPage=0
            monPdf.add_page()
        
        #Le nombre de produit de la conbinaison actuel pour la dernière page
        if j==len(DsImpressionData):
            monPdf.cell(200 ,5, txt="------------------- {CompteurProduit} visuel {Visuel1} -------------------".format(CompteurProduit=ListecompteurProduit[compteurPageProduit],Visuel1=ligne[5]), ln=1, align="C")
     
    monPdf.output("{CheminPdf}/{nomDuFichier}.pdf".format(nomDuFichier=textboxfile.get(),CheminPdf=CheminPdf))

def EcrirePdfTrainDe2(locExcel):
    tmpVisuel1=""
    tmpVisuel2=""
    DsImpressionData=TrierExcel(locExcel)
    monPdf = FPDF()
    monPdf.add_page()
    monPdf.set_font("Arial", size=20)
    monPdf.cell(200 ,5, txt="Liste recapitulative de toutes les combinaisons de visuels", ln=1, align="C") 
    monPdf.cell(200 ,10, txt="", ln=2, align="C") 
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
            monPdf.cell(200 ,5, txt="------------------- {CompteurProduit} visuel {Visuel1}/{Visuel2} -------------------".format(CompteurProduit=compteurProduit,Visuel1=tmpVisuel1,Visuel2=tmpVisuel2), ln=1, align="C")  
            compteurProduit=0

        tmpVisuel1=produit[5]
        tmpVisuel2=produit[7]
        compteurProduit+=1
        i+=1
        if i==len(DsImpressionData):
            ListecompteurProduit.append(compteurProduit)  
            monPdf.cell(200 ,5, txt="------------------- {CompteurProduit} visuel {Visuel1}/{Visuel2} -------------------".format(CompteurProduit=compteurProduit,Visuel1=produit[5],Visuel2=produit[7]), ln=1, align="C")  
            monPdf.add_page()
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
     
    monPdf.output("{CheminPdf}/{nomDuFichier}.pdf".format(nomDuFichier=textboxfile.get(),CheminPdf=CheminPdf))

def EcrirePdfTrainDe3(locExcel):
    tmpVisuel1=""
    tmpVisuel2=""
    tmpVisuel3=""
    DsImpressionData=TrierExcel(locExcel)
    monPdf = FPDF()
    monPdf.add_page()
    monPdf.set_font("Arial", size=20)
    monPdf.cell(200 ,5, txt="Liste recapitulative de toutes les combinaisons de visuels", ln=1, align="C") 
    monPdf.cell(200 ,10, txt="", ln=2, align="C") 
    monPdf.set_font("Arial", size=10)
    compteurPage=0
    compteurPageProduit=0

    compteurProduit=0
    ListecompteurProduit=[]
    i=0
    for produit in DsImpressionData:
        if (tmpVisuel1!=produit[5] or tmpVisuel2!=produit[7] or tmpVisuel3!=produit[9]) and i!=0:
            ListecompteurProduit.append(compteurProduit) 
            monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}/{Visuel3}-------------------".format(CompteurProduit=compteurProduit,Visuel1=tmpVisuel1,Visuel2=tmpVisuel2,Visuel3=tmpVisuel3), ln=1, align="C")
            compteurProduit=0

        tmpVisuel1=produit[5]
        tmpVisuel2=produit[7]
        tmpVisuel3=produit[9]
        compteurProduit+=1
        i+=1
        if i==len(DsImpressionData):
            ListecompteurProduit.append(compteurProduit)
            monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}/{Visuel3}-------------------".format(CompteurProduit=compteurProduit,Visuel1=produit[5],Visuel2=produit[7],Visuel3=produit[9]), ln=1, align="C")
            monPdf.add_page()
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
     
    monPdf.output("{CheminPdf}/{nomDuFichier}.pdf".format(nomDuFichier=textboxfile.get(),CheminPdf=CheminPdf))

def EcrirePdfTrainDe4(locExcel):
    tmpVisuel1=""
    tmpVisuel2=""
    tmpVisuel3=""
    tmpVisuel4=""
    DsImpressionData=TrierExcel(locExcel)
    monPdf = FPDF()
    monPdf.add_page()
    monPdf.set_font("Arial", size=20)
    monPdf.cell(200 ,5, txt="Liste recapitulative de toutes les combinaisons de visuels", ln=1, align="C") 
    monPdf.cell(200 ,10, txt="", ln=2, align="C") 
    monPdf.set_font("Arial", size=10)
    compteurPage=0
    compteurPageProduit=0

    compteurProduit=0
    ListecompteurProduit=[]
    i=0
    for produit in DsImpressionData:
        if (tmpVisuel1!=produit[5] or tmpVisuel2!=produit[7] or tmpVisuel3!=produit[9] or tmpVisuel4!=produit[11]) and i!=0:
            ListecompteurProduit.append(compteurProduit)  
            monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}/{Visuel3}/{Visuel4}-------------------".format(CompteurProduit=compteurProduit,Visuel1=tmpVisuel1,Visuel2=tmpVisuel2,Visuel3=tmpVisuel3,Visuel4=tmpVisuel4), ln=1, align="C")
            compteurProduit=0

        tmpVisuel1=produit[5]
        tmpVisuel2=produit[7]
        tmpVisuel3=produit[9]
        tmpVisuel4=produit[11]
        compteurProduit+=1
        i+=1
        if i==len(DsImpressionData):
            ListecompteurProduit.append(compteurProduit) 
            monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}/{Visuel3}/{Visuel4}-------------------".format(CompteurProduit=compteurProduit,Visuel1=produit[5],Visuel2=produit[7],Visuel3=produit[9],Visuel4=produit[11]), ln=1, align="C")
            monPdf.add_page() 
    j=0

    for ligne in DsImpressionData:
        
        if (tmpVisuel1!=ligne[5] or tmpVisuel2!=ligne[7] or tmpVisuel3!=ligne[9] or tmpVisuel4!=produit[11]) and j!=0:
            monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}/{Visuel3}/{Visuel4}-------------------".format(CompteurProduit=ListecompteurProduit[compteurPageProduit],Visuel1=tmpVisuel1,Visuel2=tmpVisuel2,Visuel3=tmpVisuel3,Visuel4=tmpVisuel4), ln=1, align="C")
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
        monPdf.cell(200, 5, txt= "Annonceur4 : {Annonceur4} Visuel4 : {Visuel4}".format(Annonceur4=ligne[10],Visuel4=ligne[11]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "", ln=1, align="L")
        compteurPage+=1
        if compteurPage==5:
            monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}/{Visuel3}/{Visuel4}-------------------".format(CompteurProduit=ListecompteurProduit[compteurPageProduit],Visuel1=ligne[5],Visuel2=ligne[7],Visuel3=ligne[9],Visuel4=ligne[11]), ln=1, align="C")
            compteurPage=0
            monPdf.add_page()

        tmpVisuel1=ligne[5]
        tmpVisuel2=ligne[7]
        tmpVisuel3=ligne[9]
        tmpVisuel4=ligne[11]
        j+=1    

        if j==len(DsImpressionData):
           monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}/{Visuel3}/{Visuel4}-------------------".format(CompteurProduit=ListecompteurProduit[compteurPageProduit],Visuel1=ligne[5],Visuel2=ligne[7],Visuel3=ligne[9],Visuel4=ligne[11]), ln=1, align="C")
     
    monPdf.output("{CheminPdf}/{nomDuFichier}.pdf".format(nomDuFichier=textboxfile.get(),CheminPdf=CheminPdf)) 

def EcrirePdfTrainDe5(locExcel):
    tmpVisuel1=""
    tmpVisuel2=""
    tmpVisuel3=""
    tmpVisuel4=""
    tmpVisuel5=""
    DsImpressionData=TrierExcel(locExcel)
    monPdf = FPDF()
    monPdf.add_page()
    monPdf.set_font("Arial", size=20)
    monPdf.cell(200 ,5, txt="Liste recapitulative de toutes les combinaisons de visuels", ln=1, align="C") 
    monPdf.cell(200 ,10, txt="", ln=2, align="C") 
    monPdf.set_font("Arial", size=10)
    compteurPage=0
    compteurPageProduit=0

    compteurProduit=0
    ListecompteurProduit=[]
    i=0
    for produit in DsImpressionData:
        if (tmpVisuel1!=produit[5] or tmpVisuel2!=produit[7] or tmpVisuel3!=produit[9] or tmpVisuel4!=produit[11] or tmpVisuel5!=produit[13]) and i!=0:
            ListecompteurProduit.append(compteurProduit)
            monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}/{Visuel3}/{Visuel4}/{Visuel5}-------------------".format(CompteurProduit=compteurProduit,Visuel1=tmpVisuel1,Visuel2=tmpVisuel2,Visuel3=tmpVisuel3,Visuel4=tmpVisuel4,Visuel5=tmpVisuel5), ln=1, align="C")
            compteurProduit=0

        tmpVisuel1=produit[5]
        tmpVisuel2=produit[7]
        tmpVisuel3=produit[9]
        tmpVisuel4=produit[11]
        tmpVisuel5=produit[13]
        compteurProduit+=1
        i+=1
        if i==len(DsImpressionData):
            ListecompteurProduit.append(compteurProduit)
            monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}/{Visuel3}/{Visuel4}/{Visuel5}-------------------".format(CompteurProduit=compteurProduit,Visuel1=produit[5],Visuel2=produit[7],Visuel3=produit[9],Visuel4=produit[11],Visuel5=produit[13]), ln=1, align="C")
            monPdf.add_page()
              
    j=0

    for ligne in DsImpressionData:
        
        if (tmpVisuel1!=ligne[5] or tmpVisuel2!=ligne[7] or tmpVisuel3!=ligne[9] or tmpVisuel4!=produit[11] or tmpVisuel5!=produit[13]) and j!=0:
            monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}/{Visuel3}/{Visuel4}/{Visuel5}-------------------".format(CompteurProduit=ListecompteurProduit[compteurPageProduit],Visuel1=tmpVisuel1,Visuel2=tmpVisuel2,Visuel3=tmpVisuel3,Visuel4=tmpVisuel4,Visuel5=tmpVisuel5), ln=1, align="C")
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
        monPdf.cell(200, 5, txt= "Annonceur4 : {Annonceur4} Visuel4 : {Visuel4}".format(Annonceur4=ligne[10],Visuel4=ligne[11]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "Annonceur5 : {Annonceur5} Visuel5 : {Visuel5}".format(Annonceur5=ligne[12],Visuel5=ligne[13]), ln=1, align="L")
        monPdf.cell(200, 5, txt= "", ln=1, align="L")
        compteurPage+=1
        if compteurPage==5:
            monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}/{Visuel3}/{Visuel4}/{Visuel5}-------------------".format(CompteurProduit=ListecompteurProduit[compteurPageProduit],Visuel1=ligne[5],Visuel2=ligne[7],Visuel3=ligne[9],Visuel4=ligne[11],Visuel5=ligne[13]), ln=1, align="C")
            compteurPage=0
            monPdf.add_page()

        tmpVisuel1=ligne[5]
        tmpVisuel2=ligne[7]
        tmpVisuel3=ligne[9]
        tmpVisuel4=ligne[11]
        tmpVisuel5=produit[13]
        j+=1    

        if j==len(DsImpressionData):
           monPdf.cell(200 ,5, txt="-------------------{CompteurProduit} visuel {Visuel1}/{Visuel2}/{Visuel3}/{Visuel4}/{Visuel5}-------------------".format(CompteurProduit=ListecompteurProduit[compteurPageProduit],Visuel1=ligne[5],Visuel2=ligne[7],Visuel3=ligne[9],Visuel4=ligne[11],Visuel5=ligne[13]), ln=1, align="C")
     
    monPdf.output("{CheminPdf}/{nomDuFichier}.pdf".format(nomDuFichier=textboxfile.get(),CheminPdf=CheminPdf)) 

def fonction_Principale(locExcel):
    nomdufichier=textboxfile.get()
    if locExcel=="":
        messagebox.showerror("Error", "Vous n'avez pas importé de fichier excel")
    
    elif CheminPdf=="":
        messagebox.showerror("Error", "Vous n'avez pas rentré de chemin de sortie")
    
    elif nomdufichier=="":
        messagebox.showerror("Error", "Vous n'avez rentré de nom pour le fichier")
    
    elif NumeroTrain.get()==1:
        EcrirePdfTrainDe1(locExcel)
        messagebox.showinfo("Succes", "Votre fichier a été exporté avec succès")
        print("Train de 1")
    elif NumeroTrain.get()==2:
        EcrirePdfTrainDe2(locExcel)
        messagebox.showinfo("Succes", "Votre fichier a été exporté avec succès")
        print("Train de 2")       
    elif NumeroTrain.get()==3:
        EcrirePdfTrainDe3(locExcel)
        messagebox.showinfo("Succes", "Votre fichier a été exporté avec succès")
        print("Train de 3")

    elif NumeroTrain.get()==4:
        EcrirePdfTrainDe4(locExcel)
        messagebox.showinfo("Succes", "Votre fichier a été exporté avec succès")
        print("Train de 4")

    elif NumeroTrain.get()==5:
        EcrirePdfTrainDe5(locExcel)
        messagebox.showinfo("Succes", "Votre fichier a été exporté avec succès")
        print("Train de 5")

    else:
        messagebox.showerror("Error", "Vous n'avez pas selectionné de train")
																								
def reset():
    ColonneTrainId.delete(0,END)
    ColonneTrainId.insert(-1, '1')
    ColonneCommune.delete(0,END)
    ColonneCommune.insert(-1, '5')
    ColonneAdresse.delete(0,END)
    ColonneAdresse.insert(-1, '6')
    ColonneAfficheur.delete(0,END)
    ColonneAfficheur.insert(-1, '10')
    ColonneAnnonceur1.delete(0,END)
    ColonneAnnonceur1.insert(-1, '14')
    ColonneVisuel1.delete(0,END)
    ColonneVisuel1.insert(-1, '21')
    ColonneAnnonceur2.delete(0,END)
    ColonneAnnonceur2.insert(-1, '15')
    ColonneVisuel2.delete(0,END)
    ColonneVisuel2.insert(-1, '24')
    ColonneAnnonceur3.delete(0,END)
    ColonneAnnonceur3.insert(-1, '16')
    ColonneVisuel3.delete(0,END)
    ColonneVisuel3.insert(-1, '27')
    ColonneAnnonceur4.delete(0,END)
    ColonneAnnonceur4.insert(-1, '17')
    ColonneVisuel4.delete(0,END)
    ColonneVisuel4.insert(-1, '30')
    ColonneAnnonceur5.delete(0,END)
    ColonneAnnonceur5.insert(-1, '18')
    ColonneVisuel5.delete(0,END)
    ColonneVisuel5.insert(-1, '33')     

window.config(background = "white")

# Create a File Explorer label
label_file_explorer = Label(window,
                            text = "Selectionner un fichier excel à exporter :",
                            width = 70, height = 4,
                            fg = "blue")

    
button_exploreExcel = ttk.Button(window,
                        text = "Importer",
                        command = BrowseExcel)

label_file_output = Label(window,
                            text = "Selectionner un Chemin de sortie :",
                            width = 50, height = 4,
                            fg = "blue")                        

button_outputPdf = ttk.Button(window,
                        text = "Dossier de destination",
                        command = OutputPdf) 

label_file_filename = Label(window,
                            text = "Veuillez choisir un nom pour le fichier pdf : ",
                            width = 50, height = 4,
                            fg = "black")  

textboxfile = ttk.Entry(window, width=25, justify='center')


label_TrainRadio = Label(window,
                            text = "Veuillez choisir un train : ",
                            width = 50, height = 4,
                            fg = "black")  

checkbox0 = ttk.Radiobutton(window, text='train de 1', variable=NumeroTrain, value=1)
checkbox1 = ttk.Radiobutton(window, text='train de 2', variable=NumeroTrain, value=2)
checkbox2 = ttk.Radiobutton(window, text='train de 3', variable=NumeroTrain, value=3)
checkbox3 = ttk.Radiobutton(window, text='train de 4', variable=NumeroTrain, value=4)
checkbox4 = ttk.Radiobutton(window, text='train de 5', variable=NumeroTrain, value=5)

label_NumeroFeuille = Label(window,
                            text = "Numero de la feuille : ",
                            width = 25 , height = 4,
                            fg = "black")  
NumeroFeuille= ttk.Entry(window, width=10, justify='center')
NumeroFeuille.insert(-1, '1')

button_Exporter = ttk.Button(window,
                    text = "Exporter",
                    command = lambda: fonction_Principale(locExcel))

label_ColonneTrainId = Label(window,
                            text = "Position train id :  ",
                            width = 15, height = 4,
                            fg = "black")  
ColonneTrainId = ttk.Entry(window, width=10, justify='center')
ColonneTrainId.insert(-1, '1')

label_ColonneCommune = Label(window,
                            text = "Position commune du panneau :  ",
                            width = 25 , height = 4,
                            fg = "black")  
ColonneCommune= ttk.Entry(window, width=10, justify='center')
ColonneCommune.insert(-1, '5')

label_ColonneAdresse = Label(window,
                            text = "Position addresse du panneau :  ",
                            width = 25 , height = 4,
                            fg = "black")  
ColonneAdresse= ttk.Entry(window, width=10, justify='center')
ColonneAdresse.insert(-1, '6')

label_ColonneAfficheur = Label(window,
                            text = "Position afficheur :  ",
                            width = 20 , height = 4,
                            fg = "black")  
ColonneAfficheur= ttk.Entry(window, width=10, justify='center')
ColonneAfficheur.insert(-1, '10')

label_ColonneAnnonceur1 = Label(window,
                            text = "Position annonceur 1 :  ",
                            width = 20 , height = 4,
                            fg = "black")  
ColonneAnnonceur1= ttk.Entry(window, width=10, justify='center')
ColonneAnnonceur1.insert(-1, '14')

label_ColonneVisuel1 = Label(window,
                            text = "Position visuel 1 :  ",
                            width = 20 , height = 4,
                            fg = "black")  
ColonneVisuel1= ttk.Entry(window, width=10, justify='center')
ColonneVisuel1.insert(-1, '21')

label_ColonneAnnonceur2 = Label(window,
                            text = "Position annonceur 2 :  ",
                            width = 20 , height = 4,
                            fg = "black")  
ColonneAnnonceur2= ttk.Entry(window, width=10, justify='center')
ColonneAnnonceur2.insert(-1, '15')

label_ColonneVisuel2 = Label(window,
                            text = "Position visuel 2 :  ",
                            width = 20 , height = 4,
                            fg = "black")  
ColonneVisuel2= ttk.Entry(window, width=10, justify='center')
ColonneVisuel2.insert(-1, '24')

label_ColonneAnnonceur3 = Label(window,
                            text = "Position annonceur 3 :  ",
                            width = 20 , height = 4,
                            fg = "black")  
ColonneAnnonceur3= ttk.Entry(window, width=10, justify='center')
ColonneAnnonceur3.insert(-1, '16')

label_ColonneVisuel3 = Label(window,
                            text = "Position visuel 3 :  ",
                            width = 20 , height = 4,
                            fg = "black")  
ColonneVisuel3= ttk.Entry(window, width=10, justify='center')
ColonneVisuel3.insert(-1, '27')

label_ColonneAnnonceur4 = Label(window,
                            text = "Position annonceur 4 :  ",
                            width = 20 , height = 4,
                            fg = "black")  
ColonneAnnonceur4= ttk.Entry(window, width=10, justify='center')
ColonneAnnonceur4.insert(-1, '17')

label_ColonneVisuel4 = Label(window,
                            text = "Position visuel 4 :  ",
                            width = 20 , height = 4,
                            fg = "black")  
ColonneVisuel4= ttk.Entry(window, width=10, justify='center')
ColonneVisuel4.insert(-1, '30')

label_ColonneAnnonceur5 = Label(window,
                            text = "Position annonceur 5 :  ",
                            width = 20 , height = 4,
                            fg = "black")  
ColonneAnnonceur5= ttk.Entry(window, width=10, justify='center')
ColonneAnnonceur5.insert(-1, '18')

label_ColonneVisuel5 = Label(window,
                            text = "Position visuel 5 :  ",
                            width = 20 , height = 4,
                            fg = "black")  
ColonneVisuel5= ttk.Entry(window, width=10, justify='center')
ColonneVisuel5.insert(-1, '33')

button_Reset = ttk.Button(window,
                    text = "Réinisialiser la configuration",
                    command = reset)
# Grid method is chosen for placing
# the widgets at respective positions
# in a table like structure by
# specifying rows and columns
label_file_explorer.grid(column = 0, row = 0)

button_exploreExcel.grid(column = 0, row = 1)

label_file_output.grid(column=0, row=2)

button_outputPdf.grid(column=0,row=3)

label_file_filename.grid(column=0,row=4)

textboxfile.grid(column=0, row=5)

label_TrainRadio.grid(column=0, row=6)

checkbox0.grid(column=0, row=7)
checkbox1.grid(column=0, row=8)
checkbox2.grid(column=0, row=9)
checkbox3.grid(column=0, row=10)
checkbox4.grid(column=0, row=11)

label_NumeroFeuille.grid(column=0,row=12)
NumeroFeuille.grid(column=0,row=13)

button_Exporter.grid(column = 0,row = 14,ipady=20, ipadx=20)

label_ColonneTrainId.grid(column=1,row=0)
ColonneTrainId.grid(column=1,row=1)

label_ColonneCommune.grid(column=1,row=2)
ColonneCommune.grid(column=1,row=3)

label_ColonneAdresse.grid(column=1, row=4)
ColonneAdresse.grid(column=1, row=5)

label_ColonneAfficheur.grid(column=1,row=6)
ColonneAfficheur.grid(column=1,row=7)

label_ColonneAnnonceur1.grid(column=1,row=8)
ColonneAnnonceur1.grid(column=1,row=9)

label_ColonneVisuel1.grid(column=2,row=8)
ColonneVisuel1.grid(column=2,row=9)

label_ColonneAnnonceur2.grid(column=1,row=10)
ColonneAnnonceur2.grid(column=1,row=11)

label_ColonneVisuel2.grid(column=2,row=10)
ColonneVisuel2.grid(column=2,row=11)

label_ColonneAnnonceur3.grid(column=1,row=12)
ColonneAnnonceur3.grid(column=1,row=13)

label_ColonneVisuel3.grid(column=2,row=12)
ColonneVisuel3.grid(column=2,row=13)

label_ColonneAnnonceur4.grid(column=1,row=14)
ColonneAnnonceur4.grid(column=1,row=15)

label_ColonneVisuel4.grid(column=2,row=14)
ColonneVisuel4.grid(column=2,row=15)

label_ColonneAnnonceur5.grid(column=1,row=16)
ColonneAnnonceur5.grid(column=1,row=17)

label_ColonneVisuel5.grid(column=2,row=16)
ColonneVisuel5.grid(column=2,row=17)

button_Reset.grid(column=2,row=1,ipady=15, ipadx=15)
# Let the window wait for any events
window.mainloop()
