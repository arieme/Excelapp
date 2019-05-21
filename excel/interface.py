
from tkinter import *
import xlrd
from xlwt import Workbook, Formula

def selected_column():
    print (var.get())




def find_excel():
    # lecture

   #r"C:\Users\dell\PycharmProjects\myfirstprojectpy\excelTP\Classeur1.xlsx"
    path = fich_source_entry.get()
    # Réouverture du classeur
    classeur = xlrd.open_workbook(path)

    # Récupération du nom de toutes les feuilles sous forme de liste
    nom_des_feuilles = classeur.sheet_names()

    # Récupération de la première feuille
    #feuille = classeur.sheet_by_name(nom_des_feuilles[0])  de la feuille "nom_des_feuilles[0]"
    feuille = classeur.sheet_by_name(feuille_source_entry.get())
    # recuperation de la feuille par index: feuille = classeur.sheet_by_index(0)

    #print("Lecture des cellules:")
    #for col in range(feuille.ncols):
       # print("cellule: {}".format(feuille.cell_value(0, col)))
    if ( var.get()==1):
      for r in range(feuille.nrows):
        for c in range(feuille.ncols):
           if (column_source_entry.get() == feuille.cell_value(r, c) ):
             # data = [[feuille.cell_value(r, c) for c in range(feuille.ncols)] for r in range(feuille.nrows)]
                spinc =int(spin_column.get())+1
                data = [feuille.cell_value(r, c)for r in range(spinc)]
    if (var.get()==0):
        for r in range(feuille.nrows):
            for c in range(feuille.ncols):
                if (column_source_entry.get() == feuille.cell_value(r, c)):
                    # data = [[feuille.cell_value(r, c) for c in range(feuille.ncols)] for r in range(feuille.nrows)]
                    data =[feuille.cell_value(r, c)for r in range(feuille.nrows)]



    if (var.get()==3):
        for r in range(feuille.nrows):
            for c in range(feuille.ncols):
                if (ligne_source_entry.get() == feuille.cell_value(r, c)):
                    # data = [[feuille.cell_value(r, c) for c in range(feuille.ncols)] for r in range(feuille.nrows)]
                    spinl = int(spin_ligne.get())+1
                    data = [feuille.cell_value(r, c) for r in range(spinl)]
    if (var.get()==4):
        for r in range(feuille.nrows):
            for c in range(feuille.ncols):
                if (ligne_source_entry.get() == feuille.cell_value(r, c)):
                    # data = [[feuille.cell_value(r, c) for c in range(feuille.ncols)] for r in range(feuille.nrows)]
                    data = [feuille.cell_value(r, c) for r in range(feuille.ncols)]





    # ecriture
    #path = r"D:\python\Classeur2.xls"
    path1 = fich_dest_entry.get()

    # On créer un "classeur"
    classeur = Workbook()
    # On ajoute une feuille au classeur
    feuille = classeur.add_sheet("Feuille1")

    # Ecrire les donnees
    #Ecrire une colonne
    if ((var.get() == 0) or (var.get() == 1)):
      for c in range(len(data)):
       # for r in range(len(data[c])):
           # print(data[c][r])
            print(data[c])
           # feuille.write(c, r, data[c][r])
            feuille.write(c, 0, data[c])

    if ((var1.get() == 3) or (var1.get() == 4 )):
       # Ecrire une ligne
      for c in range(len(data)):
        print(data[c])
        feuille.write(0, c, data[c])


    # Ecriture du classeur sur le disque
    classeur.save(path1)
def tout_fichier():
    #lecture
    path = fich_source_entry.get()
    classeur = xlrd.open_workbook(path)
    nom_des_feuilles = classeur.sheet_names()
    for i in range(len(nom_des_feuilles)):
      feuille = classeur.sheet_by_name(nom_des_feuilles[0])
      print(type(feuille))
      # ecriture
      path1 = fich_dest_entry.get()

      # On créer un "classeur"
      classeur = Workbook()
      # On ajoute une feuille au classeur
      feuille1 = classeur.add_sheet("Feuille {}".format(i))
      for r in range(feuille.nrows):
          for c in range(feuille.ncols):
              data = [[feuille.cell_value(r, c) for c in range(feuille.ncols)] for r in range(feuille.nrows)]

      # Ecrire les donnees
      for c in range(len(data)):
         for r in range(len(data[c])):
          feuille1.write(c, r, data[c][r])




  # Ecriture du classeur sur le disque
    classeur.save(path1)



mainapp = Tk()
var = IntVar()   #You cannot create an instance of StringVar until after the root window has been created.
var1 = IntVar()

mainapp.title("Gérer ton Excel !!")
frame = Frame(mainapp, bg="#41B77F", bd=0)

#mainapp.resizable(width = False, height=True)

#mainapp.positionfrom("user")
#geometry("XxY+40+40") 40 est le dacalage entre la borudure
screen_x = int(mainapp.winfo_screenwidth())
screen_y = int(mainapp.winfo_screenheight())
window_x = 800
window_y = 600


posX = (screen_x // 2) - (window_x // 2)
posY = (screen_y // 2) - (window_y // 2)
geo = "{}x{}+{}+{}".format(window_x, window_y,  posX, posY)
mainapp.geometry(geo)
mainapp.maxsize(800, 600)
mainapp.minsize(800, 600)
mainapp.iconbitmap("Logo_Microsoft_Excel_2013.ico")
mainapp.config(bg ='#41B77F')
right_frame= Frame(frame, bg="#41B77F")
right_frame.grid(row=0, column=1)

#label_title = Label(right_frame, text="Bienvenue sur l'application ", font=("Courrier",20), bg="#41B77F", fg="white")
#label_title.pack(padx=0, pady=0)


label_feuille_source =Label(right_frame, font=("Helvatica", 20), fg="white", bg="#41B77F", text="Nom de la feuille")
label_feuille_source.pack()
feuille_source_entry = Entry(right_frame, text="feuille source", font=("Helvatica", 20), fg="white", bg="#41B77F")
feuille_source_entry.pack()


label_column_source =Label(right_frame, font=("Helvatica", 20), fg="white", bg="#41B77F", text="Nom de la colonne")
label_column_source.pack()
column_source_entry = Entry(right_frame, font=("Helvatica", 20), fg="white", bg="#41B77F")
column_source_entry.pack()
column_radio = Radiobutton(right_frame, text="Toute  la colonne", value=0, bg="#41B77F", variable=var, command=selected_column)
column_radio.pack()
column_radio1 = Radiobutton(right_frame, text="Quelques éléments", value=1, bg="#41B77F", variable= var)
column_radio1.pack()
spin_column = Spinbox(right_frame, from_=1, to=1000)
spin_column.pack()
button_column = Button(right_frame, text = "Done", bg="white", fg="#41B77F", bd=0, command=find_excel)
button_column.pack()

label_ligne_source =Label(right_frame, font=("Helvatica", 20), fg="white", bg="#41B77F", text="Nom de la ligne")
label_ligne_source.pack()
ligne_source_entry = Entry(right_frame, font=("Helvatica", 20), fg="white", bg="#41B77F")
ligne_source_entry.pack()
ligne_radio = Radiobutton(right_frame, text="Toute  la  ligne", value=3, bg="#41B77F", variable=var1)
ligne_radio.pack()
ligne_radio1 = Radiobutton(right_frame, text="Quelques éléments", value=4, bg="#41B77F", variable=var1)
ligne_radio1.pack()
spin_ligne= Spinbox(right_frame, from_=1, to=1000)
spin_ligne.pack()
button_ligne = Button(right_frame, text = "Done", bg="white", fg="#41B77F", bd=0, command=find_excel)
button_ligne.pack()







width_image =280
height_image =300
image = PhotoImage(file="Logo_Microsoft_Excel_2013.png").zoom(35).subsample(32)
canvas = Canvas(frame, width=width_image, height=height_image, bg="#41B77F", bd=0, highlightthickness=0)
canvas.create_image(width_image/2, height_image/2, image=image)
canvas.grid(row=0, column=0, padx=80, pady=0)
left_frame = Frame(frame, bg="#41B77F")
left_frame.grid(row=1, column=0)

label_fich_source =Label(left_frame, font=("Helvatica", 20), fg="white", bg="#41B77F", text="Chemin de fihier source")
label_fich_source.pack()
fich_source_entry = Entry(left_frame, text="fich source", font=("Helvatica", 20), fg="white", bg="#41B77F")
fich_source_entry.pack()

label_fich_dest = Label(left_frame, font=("Helvatica", 20), fg="white", bg="#41B77F", text="Chemin de nouveau fichier ")
label_fich_dest.pack()
fich_dest_entry = Entry(left_frame, font=("Helvatica", 20), fg="white", bg="#41B77F")
fich_dest_entry.pack()
tout_fich_button = Button(right_frame, font=("Helvatica", 20), fg="white", bg="#41B77F", text="Copier tout le fichier", command= tout_fichier)
#tout_fich_button.pack()


frame.pack()

mainapp.mainloop()
