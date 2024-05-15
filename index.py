import tkinter as tk
from tkinter import messagebox
import openpyxl # type: ignore

def valider(event=None):
    # Récupérer les valeurs saisies dans le formulaire
    designation = champ_designation.get()
    quantite = champ_quantite.get()

    # Vérifier si les champs sont vides
    if not designation or not quantite:
        messagebox.showerror("Erreur", "Veuillez remplir tous les champs.")
        return

    # Écrire les valeurs dans le fichier Excel
    classeur = openpyxl.load_workbook("formulaire.xlsx")
    feuille = classeur.active
    feuille.append([designation, quantite])
    classeur.save("formulaire.xlsx")

    # Effacer les champs après validation
    champ_designation.delete(0, tk.END)
    champ_quantite.delete(0, tk.END)

    # Donner le focus au champ de saisie de désignation
    champ_designation.focus_set()

# Créer la fenêtre principale
fenetre = tk.Tk()
fenetre.title("Formulaire")

# Empêcher le redimensionnement de la fenêtre
fenetre.resizable(width=False, height=False)

# Créer les champs de saisie et les étiquettes pour le formulaire
etiquette_designation = tk.Label(fenetre, text="Designation:")
etiquette_designation.grid(row=0, column=0, padx=10, pady=5)
champ_designation = tk.Entry(fenetre)
champ_designation.grid(row=0, column=1, padx=10, pady=5)

etiquette_quantite = tk.Label(fenetre, text="Quantité:")
etiquette_quantite.grid(row=1, column=0, padx=10, pady=5)
champ_quantite = tk.Entry(fenetre)
champ_quantite.grid(row=1, column=1, padx=10, pady=5)

# Créer le bouton de validation
bouton_valider = tk.Button(fenetre, text="Valider", command=valider)
bouton_valider.grid(row=2, column=0, columnspan=2, pady=10)

# Lier l'événement de pression de la touche Enter à la fonction de validation
fenetre.bind('<Return>', valider)

# Lancer la boucle principale
fenetre.mainloop()
