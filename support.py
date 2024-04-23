import tkinter as tk
from tkinter import messagebox as mb
from tkinter import filedialog as fd
from PDK import substance_to_PDK , all_elements
from search import AutocompleteEntry
from tkinter import ttk
import re

refer = tk.Tkinter()
refer.title("Справка")
w = refer.winfo_screenwidth() 
h = refer.winfo_screenheight() 
w = w//2 
h = h//2 
w = w - 360
h = h - 250
refer.geometry('720x500+{}+{}'.format(w, h))
refer.resizable(False, False)
container = ttk.Frame(refer)
canvas = tk.Canvas(container, width = 700, height = 500)
scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
scrollable_frame = ttk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")
    )
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

canvas.configure(yscrollcommand=scrollbar.set)

table_6_image = tk.PhotoImage(file = "images\\table6.gif")
table_6_label = ttk.Label(scrollable_frame,image = table_6_image, width = 10)
table_6_label.pack()
table_6_label.image = table_6_image

table_8_image = tk.PhotoImage(file = "images\\table8.gif")
table_8_label = ttk.Label(scrollable_frame,image = table_8_image, width = 10)
table_8_label.pack()
table_8_label.image = table_8_image

table_9_image = tk.PhotoImage(file = "images\\table9.gif")
table_9_label = ttk.Label(scrollable_frame,image = table_9_image, width = 10)
table_9_label.pack()
table_9_label.image = table_9_image    

table_18_image = tk.PhotoImage(file = "images\\table18.gif")
table_18_label = ttk.Label(scrollable_frame,image = table_18_image, width = 10)
table_18_label.pack()
table_18_label.image = table_18_image

table_21_image = tk.PhotoImage(file = "images\\table21.gif")
table_21_label = ttk.Label(scrollable_frame,image = table_21_image, width = 10)
table_21_label.pack()
table_21_label.image = table_21_image

container.pack()
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")