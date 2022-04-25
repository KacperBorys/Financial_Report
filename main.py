
#Importujemy biblioteki i konkretne funkcje w nich zawarte
from math import ceil
from tkinter.font import BOLD
import pandas as pd
import tkinter as tk
from tkinter import CENTER, NO, RIGHT, Y, Scrollbar, ttk
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure
import numpy as np



#Importujemy potrzebne zmienne ze sprawozdania finansowego
potrzebne_zmienne = ["Zysk/(strata) netto", "AKTYWA OBROTOWE", "AKTYWA RAZEM", "Przychody ze sprzedaży",
                     "KAPITAŁ WŁASNY", "AKTYWA OBROTOWE", "ZOBOWIĄZANIA KRÓTKOTERMINOWE", "Należności handlowe",
                     "Koszty sprzedanych produktów, usług, towarów i materiałów",
                     "Zapasy", "ZOBOWIĄZANIA DŁUGOTERMINOWE"]

#Tworzymy słowniki, w których będziemy przechowywać wartości w konkretnych latach
wartosci2020 = {}
wartosci2019 = {}
wartosci2018 = {}
wartosci2017 = {}

# Wczytujemy plik Excel
file = pd.read_excel("Szablon-projektu-na-Pythona.xlsx")

#Przypisujemy zmiennym konkretne kolumny z pliku Excel
Wskazniki = file[file.columns[0]]
rok2020 = file[file.columns[1]]
rok2019 = file[file.columns[2]]
rok2018 = file[file.columns[3]]
rok2017 = file[file.columns[4]]

#Tworzymy pętlę, dzięki któej przypiszemy wartości potrzebnym nam zmiennym
i = 0
for wiersz in Wskazniki:
    for header in potrzebne_zmienne:
        if header in str(wiersz):
            wartosci2020[header] = rok2020[i]
            wartosci2019[header] = rok2019[i]
            wartosci2018[header] = rok2018[i]
            wartosci2017[header] = rok2017[i]
    i += 1


#Piszemy funkcję, dzięki któej będziemy w stanie obliczyć konkretne wskaźniki z danych lat
def calc(wartosci_z_danego_roku, wartosci_z_roku_poprzedniego):
    wskazniki = {}
    wskazniki["ROA"] = round(wartosci_z_danego_roku["Zysk/(strata) netto"] / (
        (wartosci_z_danego_roku["AKTYWA RAZEM"] + wartosci_z_roku_poprzedniego["AKTYWA RAZEM"]) / 2), 2)
    wskazniki["ROE"] = round(wartosci_z_danego_roku["Zysk/(strata) netto"]/(
        (wartosci_z_danego_roku["KAPITAŁ WŁASNY"] + wartosci_z_roku_poprzedniego["KAPITAŁ WŁASNY"]) / 2), 2)
    wskazniki["wsk_plynnosci_biezacej"] = round(wartosci_z_danego_roku["AKTYWA OBROTOWE"]/(
        (wartosci_z_danego_roku["ZOBOWIĄZANIA KRÓTKOTERMINOWE"])), 2)
    wskazniki["wsk_plynnosci_szybkiej"] = round((wartosci_z_danego_roku["AKTYWA OBROTOWE"] - wartosci_z_danego_roku["Zapasy"])/(
        wartosci_z_danego_roku["ZOBOWIĄZANIA KRÓTKOTERMINOWE"]), 2)
    wskazniki["wsk_natychmiastowej_plynnosci"] = round(wartosci_z_danego_roku["Należności handlowe"] / (
        wartosci_z_danego_roku["ZOBOWIĄZANIA KRÓTKOTERMINOWE"]), 2)
    wskazniki["wsk_ogolnego_zadluzenia"] = round((
        (wartosci_z_danego_roku["ZOBOWIĄZANIA DŁUGOTERMINOWE"] + wartosci_z_danego_roku["ZOBOWIĄZANIA KRÓTKOTERMINOWE"]) / wartosci_z_danego_roku["AKTYWA RAZEM"]), 2)
    wskazniki["wsk_pokrycia_aktywow_kapitalami_wlasnymi"] = round(
        (wartosci_z_danego_roku["KAPITAŁ WŁASNY"] / wartosci_z_danego_roku["AKTYWA RAZEM"]), 2)
    wskazniki["ROS"] = round((wartosci_z_danego_roku["Zysk/(strata) netto"] /
                             wartosci_z_danego_roku["Przychody ze sprzedaży"]), 2)
    return wskazniki

#Obliczamy wskaźniki w konkretnych latach
wskazniki_2020 = calc(wartosci2020, wartosci2019)
wskazniki_2019 = calc(wartosci2019, wartosci2018)
wskazniki_2018 = calc(wartosci2018, wartosci2017)

#Tworzymy okienko
root = tk.Tk()
root.title('Sprawozdanie finansowe')
root.configure(background="#596")

#Tabelka
style = ttk.Style()
#Konfiguracja stylu tesktu nagłówkego
style.configure("Treeview.Heading", font=(None, 14))
#Konfiguracja stylu tekstu
style.configure("Treeview", font=(None, 12), rowheight=19)

#Kolumny
tree = ttk.Treeview(root, show='headings')

tree['columns'] = ('Wskaźnik', '2020', '2019', '2018')
tree.column("Wskaźnik", anchor=CENTER, width=330)
tree.column("2020", anchor=CENTER, width=70)
tree.column("2019", anchor=CENTER, width=70)
tree.column("2018", anchor=CENTER, width=70)

#Definicja nagłowek
tree.heading('Wskaźnik', text='Wskaźnik')
tree.heading('2020', text='2020')
tree.heading('2019', text='2019')
tree.heading('2018', text='2018')

#Generujemy wartości wskaźników dla poszczególnych lat
wskazniki = []
for key in wskazniki_2020:
    wskazniki.append((f'{key}', f'{wskazniki_2020[key]}', f'{wskazniki_2019[key]}', f'{wskazniki_2018[key]}',))
    print(wskazniki)
#Dodajemy wartości wskaźników z poszzegolnych lat do tabelki
for wsk in wskazniki:
    tree.insert('', tk.END, values=wsk)
#Ustalamy miesjce naszej tabeli w oknie
tree.grid(row=0, column=1, sticky='nsew')

#Wykres1
t_1 = [wartosci2017["Zysk/(strata) netto"], wartosci2018["Zysk/(strata) netto"],
     wartosci2019["Zysk/(strata) netto"], wartosci2020["Zysk/(strata) netto"]]

fig = Figure(figsize=(5, 2), dpi=85)
fig.add_subplot(111).plot(["2017", "2018", "2019", "2020"], t_1)
fig.suptitle('Zysk/(strata) netto', fontsize=12)

for x, y in zip(["2017", "2018", "2019", "2020"], t_1):

    label = str(float("{:.2f}".format(y)) / 1000) + " tys."

    fig.gca().annotate(label,  #To jesty tekst
                       (x, y),  #To są współrzędne do umieszczenia etykiety
                       textcoords="offset points",  #Jak pozycjonowac tekst
                       xytext=(0, 9),  #Odległość od tekstu do punktów (x,y)
                       ha='center')

canvas = FigureCanvasTkAgg(fig, master=root)
canvas.draw()
canvas.get_tk_widget().grid(row=1, column=0)
current_values = fig.gca().get_yticks()
fig.gca().set_yticklabels(['{:.0f}'.format(x) for x in current_values])

#Wykres2
t_3 = [wartosci2017["Koszty sprzedanych produktów, usług, towarów i materiałów"], wartosci2018["Koszty sprzedanych produktów, usług, towarów i materiałów"],
       wartosci2019["Koszty sprzedanych produktów, usług, towarów i materiałów"], wartosci2020["Koszty sprzedanych produktów, usług, towarów i materiałów"]]

fig2 = Figure(figsize=(6, 2), dpi=85)
fig2.add_subplot(111).plot(["2017", "2018", "2019", "2020"], t_3)
fig2.suptitle('Koszty sprzedanych produktów, usług, towarów i materiałów', fontsize=12)

for x, y in zip(["2017", "2018", "2019", "2020"], t_3):

    label = str(float("{:.2f}".format(y)) / 1000) + " tys."

    fig2.gca().annotate(label,  #To jesty tekst
                       (x, y),  #To są współrzędne do umieszczenia etykiety
                       textcoords="offset points",  #Jak pozycjonowac tekst
                       xytext=(0, -7),  #Odległość od tekstu do punktów (x,y)
                       ha='center')

canvas = FigureCanvasTkAgg(fig2, master=root)
canvas.draw()
canvas.get_tk_widget().grid(row=1, column=1)
current_values = fig2.gca().get_yticks()
fig2.gca().set_yticklabels(['{:.0f}'.format(x) for x in current_values])

#Wykres3
t_2 = [wartosci2017["Przychody ze sprzedaży"], wartosci2018["Przychody ze sprzedaży"],
       wartosci2019["Przychody ze sprzedaży"], wartosci2020["Przychody ze sprzedaży"]]

fig1 = Figure(figsize=(5, 2), dpi=85)
fig1.add_subplot(111).plot(["2017", "2018", "2019", "2020"], t_2)
fig1.suptitle('Przychody ze sprzedaży', fontsize=12)

for x, y in zip(["2017", "2018", "2019", "2020"], t_2):

    label = str(float("{:.2f}".format(y)) / 1000) + " tys."

    fig1.gca().annotate(label,  #To jesty tekst
                       (x, y),  #To są współrzędne do umieszczenia etykiety
                       textcoords="offset points",  #Jak pozycjonowac tekst
                       xytext=(0, 5),  #Odległość od tekstu do punktów (x,y)
                       ha='center')

canvas = FigureCanvasTkAgg(fig1, master=root)
canvas.draw()
canvas.get_tk_widget().grid(row=1, column=2)
current_values = fig1.gca().get_yticks()
fig1.gca().set_yticklabels(['{:.0f}'.format(x) for x in current_values])


#Tworzymy wykres Du-ponta
#Rok 2020
label2020 = tk.Label(root, text="2020", background="#00A8E8", foreground="#fff", font=(None, 16))
label2020.grid(row=5, column=2, padx=18, pady=12)

wskps2020_label = tk.Label(root, text=f'Wskaznik plynnosci szybkiej = {wskazniki_2020["wsk_plynnosci_szybkiej"]}', background="#00A8E8", foreground="#fff", font=(None, 16))
wskps2020_label.grid(row=6, column=2, padx=18, pady=12)

wskpb2020_label = tk.Label(root, text=f'Wskaznik plynnosci biezacej = {wskazniki_2020["wsk_plynnosci_biezacej"]}', background="#00A8E8", foreground="#fff", font=(None, 16))
wskpb2020_label.grid(row=7, column=2, padx=18, pady=12)

roe2020_label = tk.Label(root, text=f'ROE = {wskazniki_2020["ROE"]}', background="#00A8E8", foreground="#fff", font=(None, 16))
roe2020_label.grid(row=8, column=2, padx=18, pady=12)

roa2020_label = tk.Label(root, text=f'ROA = {wskazniki_2020["ROA"]}', background="#00A8E8", foreground="#fff", font=(None, 16))
roa2020_label.grid(row=9, column=2, padx=18, pady=12)

ros2020_label = tk.Label(root, text=f'ROS = {wskazniki_2020["ROS"]}', background="#00A8E8", foreground="#fff", font=(None, 16))
ros2020_label.grid(row=10, column=2, padx=18, pady=12)

#Rok 2019
label2019 = tk.Label(root, text="2019", background="#95eb34", foreground="#fff", font=(None, 16))
label2019.grid(row=5, column=1, padx=18, pady=12)

wskps2019_label = tk.Label(root, text=f'Wskaznik plynnosci szybkiej = {wskazniki_2019["wsk_plynnosci_szybkiej"]}', background="#95eb34", foreground="#fff", font=(None, 16))
wskps2019_label.grid(row=6, column=1, padx=18, pady=12)

wskpb2019_label = tk.Label(root, text=f'Wskaznik plynnosci biezacej = {wskazniki_2019["wsk_plynnosci_biezacej"]}', background="#95eb34", foreground="#fff", font=(None, 16))
wskpb2019_label.grid(row=7, column=1, padx=18, pady=12)

roe2019_label = tk.Label(root, text=f'ROE = {wskazniki_2019["ROE"]}', background="#95eb34", foreground="#fff", font=(None, 16))
roe2019_label.grid(row=8, column=1, padx=18, pady=12)

roa2019_label = tk.Label(root, text=f'ROA = {wskazniki_2019["ROA"]}', background="#95eb34", foreground="#fff", font=(None, 16))
roa2019_label.grid(row=9, column=1, padx=18, pady=12)

ros2019_label = tk.Label(root, text=f'ROS = {wskazniki_2019["ROS"]}', background="#95eb34", foreground="#fff", font=(None, 16))
ros2019_label.grid(row=10, column=1, padx=18, pady=12)

#Rok 2018
label2018 = tk.Label(root, text="2018", background="#eb4034", foreground="#fff", font=(None, 16))
label2018.grid(row=5, column=0, padx=18, pady=12)

wskps2018_label = tk.Label(root, text=f'Wskaznik plynnosci szybkiej = {wskazniki_2018["wsk_plynnosci_szybkiej"]}', background="#eb4034", foreground="#fff", font=(None, 16))
wskps2018_label.grid(row=6, column=0, padx=18, pady=12)

wskpb2018_label = tk.Label(root, text=f'Wskaznik plynnosci biezacej = {wskazniki_2018["wsk_plynnosci_biezacej"]}', background="#eb4034", foreground="#fff", font=(None, 16))
wskpb2018_label.grid(row=7, column=0, padx=18, pady=12)

roe2018_label = tk.Label(root, text=f'ROE = {wskazniki_2018["ROE"]}', background="#eb4034", foreground="#fff", font=(None, 16))
roe2018_label.grid(row=8, column=0, padx=18, pady=12)

roa2018_label = tk.Label(root, text=f'ROA = {wskazniki_2018["ROA"]}', background="#eb4034", foreground="#fff", font=(None, 16))
roa2018_label.grid(row=9, column=0, padx=18, pady=12)

ros2018_label = tk.Label(root, text=f'ROS = {wskazniki_2018["ROS"]}', background="#eb4034", foreground="#fff", font=(None, 16))
ros2018_label.grid(row=10, column=0, padx=18, pady=12)


#Rozpocznij działanie aplikacji
root.mainloop()