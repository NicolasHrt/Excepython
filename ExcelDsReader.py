# Reading an excel file using Python
from ast import Num
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
window.title('File Explorer')

# Set window size
window.geometry("410x500")

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
     
    monPdf.output("{CheminPdf}/{nomDuFichier}.pdf".format(nomDuFichier=textboxfile.get(),CheminPdf=CheminPdf))



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
     
    monPdf.output("{CheminPdf}/{nomDuFichier}.pdf".format(nomDuFichier=textboxfile.get(),CheminPdf=CheminPdf))


def EcrirePdfTrainDe4(locExcel):
    tmpVisuel1=""
    tmpVisuel2=""
    tmpVisuel3=""
    tmpVisuel4=""
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
        if (tmpVisuel1!=produit[5] or tmpVisuel2!=produit[7] or tmpVisuel3!=produit[9] or tmpVisuel4!=produit[11]) and i!=0:
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

def fonction_Principale(locExcel):
    nomdufichier=textboxfile.get()
    if locExcel=="":
        messagebox.showerror("Error", "Vous n'avez pas importé de fichier excel")
    
    elif CheminPdf=="":
        messagebox.showerror("Error", "Vous n'avez pas rentré de chemin de sortie")
    
    elif nomdufichier=="":
        messagebox.showerror("Error", "Vous n'avez rentré de nom pour le fichier")
    

    if NumeroTrain.get()==2:
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
        messagebox.showinfo("Succes", "Votre fichier a été exporté avec succès")
        print("Train de 5")

    else:
        messagebox.showerror("Error", "Vous n'avez pas selectionné de train")
																								





window.config(background = "white")

# Create a File Explorer label
label_file_explorer = Label(window,
                            text = "Selectionner un fichier excel à exporter :",
                            width = 50, height = 4,
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

checkbox1 = ttk.Radiobutton(window, text='train de 2', variable=NumeroTrain, value=2)
checkbox2 = ttk.Radiobutton(window, text='train de 3', variable=NumeroTrain, value=3)
checkbox3 = ttk.Radiobutton(window, text='train de 4', variable=NumeroTrain, value=4)
checkbox4 = ttk.Radiobutton(window, text='train de 5', variable=NumeroTrain, value=5)

button_Exporter = ttk.Button(window,
                    text = "Exporter",
                    command = lambda: fonction_Principale(locExcel))

# Grid method is chosen for placing
# the widgets at respective positions
# in a table like structure by
# specifying rows and columns
label_file_explorer.grid(column = 0, row = 1)

button_exploreExcel.grid(column = 0, row = 2)

label_file_output.grid(column=0, row=3)

button_outputPdf.grid(column=0,row=4)

label_file_filename.grid(column=0,row=5)

textboxfile.grid(column=0, row=6)

label_TrainRadio.grid(column=0, row=7)

checkbox1.grid(column=0, row=8)
checkbox2.grid(column=0, row=9)
checkbox3.grid(column=0, row=10)
checkbox4.grid(column=0, row=11)

button_Exporter.grid(column = 0,row = 13)


# Let the window wait for any events
window.mainloop()
