import datetime
import tkinter as tk
from tkinter import CENTER, ttk
from tkinter import messagebox
from docxtpl import DocxTemplate

affectation_list = ["affectation1", "affectation2", "affectation3", "affectation4", "affectation5"]

def clear():
    quantity_entry.delete(0, tk.END)
    quantity_entry.insert(0, "0.0")
    designation_entry.delete(0, tk.END)
    designation_entry.insert(0, "Designation")
    num_serie_entry.delete(0, tk.END)
    num_serie_entry.insert(0, "Num Serie")
    remarque_entry.delete(0, tk.END)
    remarque_entry.insert(0, "Remarque")

def clear_all():
    clear()
    first_name_entry.delete(0, tk.END)
    first_name_entry.insert(0, "First Name")
    last_name_entry.delete(0, tk.END)
    last_name_entry.insert(0, "Last Name")
    matricule_entry.delete(0, tk.END)
    matricule_entry.insert(0, "Matricule")
    affectation_combobox.delete(0, tk.END)
    affectation_combobox.insert(0, "affctation 1")
    motifs_entry.delete(0, tk.END)
    motifs_entry.insert(0, "Motifs")
    

# REPRESENTS ALL THE LINES IN THE TREE VIEW
invoice_list = []
def add_item():
    qty = quantity_entry.get()
    designation = designation_entry.get()
    num_serie = num_serie_entry.get()
    remarque = remarque_entry.get()
    invoice_item = [qty, designation, num_serie, remarque]
    treeview.insert('', 0, values = invoice_item)
    clear()
    invoice_list.append(invoice_item) # THIS WILL CREATE A NESTED LIST LIKE THE ONE IN doc_gen.py

def new_invoice():
    first_name_entry.delete(0, tk.END)
    first_name_entry.insert(0, "First Name")
    last_name_entry.delete(0, tk.END)
    last_name_entry.insert(0, "Last Name")
    matricule_entry.delete(0, tk.END)
    matricule_entry.insert(0, "matricule Number")
    motifs_entry.delete(0, tk.END)
    motifs_entry.insert(0, "Motifs")
    quantity_entry.delete(0, tk.END)
    quantity_entry.insert(0, 1)
    affectation_combobox.set(affectation_list[0])
    designation_entry.delete(0, tk.END)
    designation_entry.insert(0, "Desgination")
    num_serie_entry.delete(0, tk.END)
    num_serie_entry.insert(0, "Numero Serie")
    remarque_entry.delete(0, tk.END)
    remarque_entry.insert(0, "Remarque")
    
    clear_all()
    treeview.delete(*treeview.get_children())

def generate_invoice():
    doc = DocxTemplate("affectation_template.docx")
    full_name = first_name_entry.get() + last_name_entry.get()
    matricule = matricule_entry.get()
    motifs = motifs_entry.get()
    affectation = affectation_combobox.get()
    num_fiche = 0
    # opening the file
    f = open('num_fiche.txt', 'r')
    # storing the number found in num_fiche
    for each in f:
        num_fiche = each
        print(num_fiche)
    # incrementing the number
    num_fiche = int(num_fiche) + 1
    print(num_fiche)
    # storing the new number in the same file
    
    f = open('num_fiche.txt', 'w')
    f.write(str(num_fiche))
    f.close()
    # RENDERING THE DOC
    doc.render({"nom_complet": full_name,
                "matricule" : matricule,
                "invoice_list": invoice_list,
                "motifs": motifs,
                "affectation": affectation,
                "date": datetime.datetime.now().strftime("%Y-%m-%d"),
                "numero_fiche": num_fiche})
    # each invoice generation we have to generate a unique doc name because the same fil will be over written
    doc_name = "BA" + "_" + full_name + "_" + affectation + ".docx"
    doc.save(doc_name)
def delete_item():
    current_item = treeview.focus()
    # print(treeview.item(current_item))
    # IT SHOULD BE DELETED FROM THE TREEVIEW AND THE INVOICE_LIST
    treeview.delete(current_item)


# declaring what we need 
root = tk.Tk()
root.resizable(False, False)
root.geometry("1100x850")
style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark") # USE THIS THEME BY DEFAULT
tree_columns = ("Qty", "Designation", "Numero Serie", "Remarque")

# DECLARING THE FRAME
frame = ttk.Frame(root)
frame.pack()

informations_agent = ttk.LabelFrame(frame, text="Informations d'agent")
informations_agent.grid(row = 0, column = 0, padx=50, pady=10)

first_name_entry = ttk.Entry(informations_agent)
first_name_entry.insert(0, "First Name")
first_name_entry.bind("<FocusIn>", lambda e: first_name_entry.delete('0', 'end'))
first_name_entry.grid(row = 0, column = 0, padx=20, pady = (0, 5), sticky="ew")

last_name_entry = ttk.Entry(informations_agent)
last_name_entry.insert(0, "Last Name")
last_name_entry.bind("<FocusIn>", lambda e: last_name_entry.delete('0', 'end'))
last_name_entry.grid(row = 0, column = 1, padx=150, pady = (0, 5), sticky="ew")


matricule_entry = ttk.Entry(informations_agent)
matricule_entry.insert(0, "Matricule")
matricule_entry.bind("<FocusIn>", lambda e: matricule_entry.delete('0', 'end'))
matricule_entry.grid(row = 0, column = 2, padx=5, pady = (0, 5), sticky="ew")


affectation_combobox = ttk.Combobox(informations_agent, values=affectation_list)
affectation_combobox.current(0) # set the first value as a its default value
affectation_combobox.grid(row = 1, column = 0, padx=20, pady = (0, 5), sticky="ew")

motifs_entry = ttk.Entry(informations_agent)
motifs_entry.insert(0, "Motifs")
motifs_entry.bind("<FocusIn>", lambda e: motifs_entry.delete('0', 'end'))
motifs_entry.grid(row = 1, column = 1,columnspan=3,  padx=150, pady = (0, 5), sticky="ew")


informations_materiel = ttk.LabelFrame(frame, text="Informations du materiel")
informations_materiel.grid(row = 3, column = 0, padx=50, pady=10)

quantity_entry = ttk.Spinbox(informations_materiel, from_=1, to=100)
quantity_entry.insert(0, "Qty")
quantity_entry.grid(row = 0, column = 0, padx=20, pady = (0, 5), sticky="ew")

designation_entry = ttk.Entry(informations_materiel)
designation_entry.insert(0, "Designation")
designation_entry.bind("<FocusIn>", lambda e: designation_entry.delete('0', 'end'))
designation_entry.grid(row = 0, column = 1, padx=150, pady = (0, 5), sticky="ew")


num_serie_entry = ttk.Entry(informations_materiel)
num_serie_entry.insert(0, "Numero Serie")
num_serie_entry.bind("<FocusIn>", lambda e: num_serie_entry.delete('0', 'end'))
num_serie_entry.grid(row = 0, column = 2, padx=5, pady = (0, 5), sticky="ew")

remarque_entry = ttk.Entry(informations_materiel)
remarque_entry.insert(0, "Remarque")
remarque_entry.bind("<FocusIn>", lambda e: remarque_entry.delete('0', 'end'))
remarque_entry.grid(row = 1, column =  0, columnspan=1,  padx=(18, 5), pady = (0, 5), sticky="nsew")

button = ttk.Button(informations_materiel, text="Add item", command = add_item)
button.grid(row = 3, column = 2, padx=(5, 10), pady=10, sticky="ew")


button = ttk.Button(informations_materiel, text="Delete item", command = delete_item)
button.grid(row = 3, column = 1, padx=(20, 5), pady=10, sticky="e")


# WE CAN USE THE GRID AND USE THE STICK TO EAST TO MAKE STICK TO THE RIGHT...

generate_invoice_button = ttk.Button(root, text="Generate Invoice", command=generate_invoice)
generate_invoice_button.place(x = 845, y = 615)

new_invoice_button = ttk.Button(root, text="New Invoice", command=new_invoice)
new_invoice_button.place(x = 730, y = 615)




treeFrame = ttk.Frame(frame)
treeFrame.grid(row = 5, column = 0, pady = 10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side = "right", fill="y")

treeview = ttk.Treeview(treeFrame, show="headings", 
                        yscrollcommand=treeScroll.set, columns=tree_columns, height=13)

treeview.column("Qty", anchor=CENTER)
treeview.heading("Qty", text="Qty")
treeview.column("Designation", anchor=CENTER)
treeview.heading("Designation", text="Designation")
treeview.column("Numero Serie", anchor=CENTER)
treeview.heading("Numero Serie", text="Numero Serie")
treeview.column("Remarque", anchor=CENTER)
treeview.heading("Remarque", text="Remarque")


# WHENEVER WE CLICK ON AN ITEM ON THE LIST, IT SHOULD RETRIEVED
# treeview.bind('<ButtonRelease-1>', selected_item)
treeview.pack()
treeScroll.config(command=treeview.yview)



""" switching between two modes (light and dark)
mode_switch = ttk.Checkbutton(
    informations_agent, text = "Mode", style="Switch", command=toggle_mode
)
mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky="nsew")
"""


root.mainloop()