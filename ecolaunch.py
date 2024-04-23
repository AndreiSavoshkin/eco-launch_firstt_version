# -*- coding: utf-8 -*-

import pandas as pd
import time
import tkinter as tk
from tkinter import messagebox as mb
from tkinter import ttk
from tkinter import filedialog as fd
from PDK import substance_to_PDK , all_elements
from nozologii import nozologii_to_disease
from search import AutocompleteEntry
import re
from PIL import Image, ImageTk
import datetime


global excel_file, n_of_elements, r, V, decrease, n_of_elements_plan, H , A , F , m , n , nu, rank_def, names_elements_plan, names_disease, n_of_disease, n_of_factory,counter_of_sub, L
L = []
names_elements = []
counter_of_sub = 0
rank_def = {}
names_elements_plan = []
n_of_disease = 0    
n_of_elements_plan = 0

A = F = m = n = nu = 0
#Appdef


def Begin():
    next_button_0.place(x = 640, y = 165)
    gen_frame.place(x = 100, y = 100)
    begin_button.place_forget()
    to_begin_button.place(x = 415,y = 560)
    begin_frame.place_forget()
    inf_label_1.place(x=16,y=325)
    inf_label_2.place(x=224,y=325)
    inf_label_3.place(x=495,y=325)
    example_1['bg'] = '#E0FFFF'
    example_2['bg'] = '#E0FFFF'
    example_3['bg'] = '#E0FFFF'
    example_4['bg'] = '#E0FFFF'
    example_5['bg'] = '#E0FFFF'
    example_6['bg'] = '#E0FFFF'
    example_1_label.place_forget()
    example_2_label.place_forget()
    example_3_label.place_forget()
    example_4_label.place_forget()
    example_5_label.place_forget()
    example_6_label.place_forget()

def _on_mousewheel(event):
    global canvas_1, canvas_2
    canvas_1.yview_scroll(-1*(event.delta//120), "units")
    canvas_2.yview_scroll(-1*(event.delta//120), "units")

def Reference():
    global canvas_1,canvas_2
    w = root.winfo_screenwidth() 
    h = root.winfo_screenheight() 
    w = w//2 
    h = h//2 
    w = w - 345
    h = h - 215
    spravka = tk.Toplevel()
    spravka.geometry('690x430+{}+{}'.format(w, h))
    spravka.resizable(False, False)
    spravka.title("Справка")
    osnova = ttk.Notebook(spravka,)
    osnova.pack()

    f1 = tk.Frame(osnova)
    osnova.add(f1, text="Как работать с программой?")
    load = Image.open("images\\f1.png")
    scrollbar = ttk.Scrollbar(f1, orient="vertical")
    canvas_1 = tk.Canvas(f1)
    canvas_1.config(width=1000, height=1000)
    scrollbar.config(command=canvas_1.yview)
    canvas_1.config(yscrollcommand=scrollbar.set)
    scrollbar.pack(side = tk.RIGHT, fill = tk.BOTH)
    canvas_1.pack(expand=tk.YES, fill=tk.BOTH)
    canvas_1.config(scrollregion=(0,0,675,545))
    canvas_1.bind_all("<MouseWheel>", _on_mousewheel)
    #Make 2nd tab
    f2 = tk.Frame(osnova)
    #Add 2nd tab 
    osnova.add(f2, text="Возможные ошибки")
    load = Image.open("images\\f2.png")
    f2_image  = ImageTk.PhotoImage(load)
    scrollbar = ttk.Scrollbar(f2, orient="vertical")
    canvas_2 = tk.Canvas(f2)
    canvas_2.config(width=1000, height=1000)
    scrollbar.config(command=canvas_2.yview)
    canvas_2.config(yscrollcommand=scrollbar.set)
    scrollbar.pack(side = tk.RIGHT, fill = tk.BOTH)
    canvas_2.pack(expand=tk.YES, fill=tk.BOTH)
    canvas_2.config(scrollregion=(0,0,675,545))
    image_f2 = canvas_2.create_image(0, 0, image=f2_image, anchor=tk.NW)
    canvas_2.bind_all("<MouseWheel>", _on_mousewheel)
    image_f2.pack()
    
    osnova.select(f1)
    osnova.enable_traversal()

def To_begin():
    global excel_file, n_of_elements,r,V, decrease, n_of_factory,counter_of_sub, L

    to_begin_button.place_forget()
    begin_button.place(x = 415,y = 560)
    gen_frame.place_forget()
    begin_frame.place(x = 0,y =  100)
    elem_Label.place_forget()
    elem_entry.place_forget()
    next_button_1.place_forget()
    next_button_2.place_forget()
    file_label['text'] = ''
    elem_entry.delete(first = 0, last = 20)
    v_entry.place_forget()
    v_label.place_forget()
    next_button_11.place_forget()
    r_entry.place_forget()
    r_label.place_forget()
    next_button_10.place_forget()
    decrease_label.place_forget()
    decrease_entry.place_forget()
    L_label.place_forget()
    L_entry.place_forget()
    next_button_5.place_forget()
    try:
        decrease_entry.delete(0,100)
    except:
        None
    factory_Label.place_forget()
    factory_entry.place_forget()
    try:
        L_entry.delete(0,100)
    except:
        None
    factory_Label.place_forget()
    factory_entry.place_forget()
    names_elements_entry[counter_of_sub].destroy()
    name_label.destroy()    
    counter_of_sub = 0
    
def exmpl_1():
    example_1['bg'] = '#1E90FF'
    example_2['bg'] = '#E0FFFF'
    example_3['bg'] = '#E0FFFF'
    example_4['bg'] = '#E0FFFF'
    example_5['bg'] = '#E0FFFF'
    example_6['bg'] = '#E0FFFF'
    example_2_label.place_forget()
    example_3_label.place_forget()
    example_4_label.place_forget()
    example_5_label.place_forget()
    example_6_label.place_forget()
    example_1_label.place(x = 5,y = 20)
def exmpl_2():
    example_1['bg'] = '#E0FFFF'
    example_2['bg'] = '#1E90FF'
    example_3['bg'] = '#E0FFFF'
    example_4['bg'] = '#E0FFFF'
    example_5['bg'] = '#E0FFFF'
    example_6['bg'] = '#E0FFFF'
    example_1_label.place_forget()
    example_3_label.place_forget()
    example_4_label.place_forget()
    example_5_label.place_forget()
    example_6_label.place_forget()
    example_2_label.place(x = 5  ,y = 20)
def exmpl_3():
    example_1['bg'] = '#E0FFFF'
    example_2['bg'] = '#E0FFFF'
    example_3['bg'] = '#1E90FF'
    example_4['bg'] = '#E0FFFF'
    example_5['bg'] = '#E0FFFF'
    example_6['bg'] = '#E0FFFF'
    example_1_label.place_forget()
    example_2_label.place_forget()
    example_4_label.place_forget()
    example_5_label.place_forget()
    example_6_label.place_forget()
    example_3_label.place(x = 5  ,y = 20)
def exmpl_4():
    example_1['bg'] = '#E0FFFF'
    example_2['bg'] = '#E0FFFF'
    example_3['bg'] = '#E0FFFF'
    example_4['bg'] = '#1E90FF'
    example_5['bg'] = '#E0FFFF'
    example_6['bg'] = '#E0FFFF'
    example_1_label.place_forget()
    example_2_label.place_forget()
    example_3_label.place_forget()
    example_5_label.place_forget()
    example_6_label.place_forget()
    example_4_label.place(x = 5  ,y = 20)
def exmpl_5():
    example_1['bg'] = '#E0FFFF'
    example_2['bg'] = '#E0FFFF'
    example_3['bg'] = '#E0FFFF'
    example_4['bg'] = '#E0FFFF'
    example_5['bg'] = '#1E90FF'
    example_6['bg'] = '#E0FFFF'
    example_1_label.place_forget()
    example_2_label.place_forget()
    example_3_label.place_forget()
    example_4_label.place_forget()
    example_6_label.place_forget()
    example_5_label.place(x = 5  ,y = 20)
def exmpl_6():
    example_1['bg'] = '#E0FFFF'
    example_2['bg'] = '#E0FFFF'
    example_3['bg'] = '#E0FFFF'
    example_4['bg'] = '#E0FFFF'
    example_5['bg'] = '#E0FFFF'
    example_6['bg'] = '#1E90FF'
    example_1_label.place_forget()
    example_2_label.place_forget()
    example_3_label.place_forget()
    example_4_label.place_forget()
    example_5_label.place_forget()
    example_6_label.place(x = 5, y = 20)

def matches(fieldValue, acListEntry):
    pattern = re.compile(re.escape(fieldValue) + '.*', re.IGNORECASE)
    return re.match(pattern, acListEntry)

def Openfile():
    try:
        file_name = fd.askopenfilename()
    except:
        mb.showerror("Ошибка", "Должен быть выбран файл")
    file_label['text'] = file_name
    
    global excel_file
    excel_file = file_name    
    next_button_1.place(x = 640, y = 315)

def Next_null():
    next_button_1.place(x = 640, y = 315)
    next_button_0.place_forget()
    file_text.place(x = 20, y = 180)
    but_file.place(x = 400, y = 216)
    file_label.place(x = 20, y = 220)
    inf_label_1.place_forget()
    inf_label_2.place_forget()
    inf_label_3.place_forget()

def Next_one():
    file_label.place_forget()
    but_file.place_forget()
    file_text.place_forget()
    next_button_1.place_forget()
    elem_Label.place(x = 20, y = 200)
    elem_entry.place(x = 400, y = 204)
    elem_entry.focus_set()
    next_button_2.place(x = 640, y = 315)
    

'''
def Get_names_elements():
    global n_of_elements,names_elements_entry,counter_of_sub,name_label
    if counter_of_sub < n_of_elements:
        str_text = 'Вещество '+str(counter_of_sub+1)+' (начинайте вводить)'
        name_label = tk.Label(text = str_text,font = 'Times 14')
        name_label.place(x = 200,y = 280)
        names_elements_entry[counter_of_sub].place(x = 200, y = 320)
        names_elements_entry[counter_of_sub].focus_set()
        '''
    
def Next_two():
    global n_of_elements, counter_of_sub,names_elements_entry, counter_of_sub
    try:
        n_of_elements = int(elem_entry.get())
        counter_of_sub = 0
        factory_Label.place(x = 20, y = 200)
        factory_entry.place(x = 400, y = 204)
        factory_entry.focus_set()
        next_button_5.place(x = 640, y = 315)        
        elem_Label.place_forget()
    except:
        mb.showerror("Ошибка", "Должно быть введено натуральное число")

def Next_five():
    global n_of_factory, L, counter_of_sub
    try:
        n_of_factory = int(factory_entry.get())
        factory_entry.delete(0,100)
        next_button_5.place_forget()
        factory_Label.place_forget()
        factory_entry.place_forget()
        counter_of_sub = 0
        next_button_10.place(x = 640, y = 315)
        r_entry.place(x = 400, y = 204)
        r_entry.focus_set()
        r_label.place(x = 20, y = 200) 
        next_button_5.place_forget()
    except:
        mb.showerror("Ошибка","Должно быть введено натуральное число")
        
def Get_r():
    global r
    try:
        r = float(r_entry.get())
        r_entry.delete(1,100)
    except:
        mb.showerror("Ошибка", "Должно быть введено число")
    

def Quit():
    root.destroy()
    
def Next_ten():
    global r
    try:
        r = float(r_entry.get())
        r_entry.place_forget()
        r_entry.delete(0,100)
        r_label.place_forget()
        v_entry.place(x = 400, y = 204)
        v_entry.focus_set()
        v_label.place(x = 20, y = 200)
        next_button_10.place_forget()
        next_button_11.place(x = 640, y = 315)

    except:
        mb.showerror("Ошибка", "Должно быть введено число") 

def Next_eleven():
    global V,n_of_disease,counter_of_sub
    try:
        V = float(v_entry.get())
        v_entry.delete(0,100)
        next_button_11.place_forget()
        v_entry.place_forget()
        v_label.place_forget()
        counter_of_sub = 0
        next_button_13.place(x = 640, y = 315)
        
    except:
        mb.showerror("Ошибка", "Должно быть введено число")
    
def Next_thirteen():
    global rank_def,counter_of_sub
    sup_blowout_entry.place(x = 500, y = 240)
    sup_blowout_entry.focus_set()
    sup_blowout_label.place(x = 20, y = 200)
    next_button_13.place_forget()
    next_button_14.place(x = 640, y = 315)
    
def Next_fourteen():
    global A
    try:
        A = float(sup_blowout_entry.get())
        sup_blowout_entry.delete(0,100)
        sup_blowout_label['text'] = 'Введите коэффициент, учитывающий скорость оседания загрязняющих веществ \n(газообразных и аэрозолей, включая твердые частицы) в атмосферном воздухе'
        next_button_14.place_forget()
        next_button_16.place(x = 640, y = 315)
    except:
        mb.showerror("Ошибка", "Должно быть введено натуральное число")


def Next_sixteen():
    global F
    try:
        F = float(sup_blowout_entry.get())
        sup_blowout_entry.delete(0,100)
        sup_blowout_label['text'] = 'Введите коэффициент, учитывающий условия выброса из устья источника выброса (m)'
        next_button_16.place_forget()
        next_button_17.place(x = 640, y = 315)
    except:
        mb.showerror("Ошибка", "Должно быть введено натуральное число")

def Next_seventeen():
    global m
    try:
        m = float(sup_blowout_entry.get())
        sup_blowout_entry.delete(0,100)
        sup_blowout_label['text'] = 'Введите коэффициент, учитывающий условия выброса из устья источника выброса (n)'
        next_button_17.place_forget()
        next_button_18.place(x = 640, y = 315)
    except:
        mb.showerror("Ошибка", "Должно быть введено натуральное число")

def Next_eightteen():
    global n
    try:
        n = float(sup_blowout_entry.get())
        sup_blowout_entry.delete(0,100)
        sup_blowout_label['text'] = 'Введите коэффициент, учитывающий влияние рельефа местности'
        next_button_18.place_forget()
        next_button_19.place(x = 640, y = 315)
    except:
        mb.showerror("Ошибка", "Должно быть введено натуральное число")

def Next_nineteen():
    global nu,counter_of_sub, n_of_elements_plan
    try:
        nu = float(sup_blowout_entry.get())
        sup_blowout_entry.delete(0,100)
        sup_blowout_entry.place_forget()
        sup_blowout_label.place_forget()
        next_button_19.place_forget()
        next_button_20.place(x = 640, y = 315)
        elem_Label.place(x = 20, y = 200)
        elem_entry.place(x = 600, y = 204)
        elem_entry.focus_set()
        elem_entry.delete(0,100)
        elem_Label['text'] = 'Введите количество веществ по плану снижения предприятиями'
        counter_of_sub = 0
    except:
        mb.showerror("Ошибка", "Должно быть введено натуральное число")

def Get_names_elements_plan():
    global n_of_elements_plan,names_elements_entry,counter_of_sub,name_label
    if counter_of_sub < n_of_elements_plan:
        str_text = 'Вещество '+str(counter_of_sub+1)+' (начинайте вводить)'
        name_label = tk.Label(text = str_text,font = 'Times 14')
        name_label.place(x = 200,y = 280)
        names_elements_entry[counter_of_sub].place(x = 200, y = 320)
        names_elements_entry[counter_of_sub].focus_set()

def Next_twenty():
    global counter_of_sub, n_of_elements_plan,names_elements_plan, names_elements_entry
    try:
        n_of_elements_plan = int(elem_entry.get())
        next_button_20.place_forget()
        autocompleteList = all_elements
        names_elements_entry = []
        names_elements_plan = ['']*n_of_elements_plan
        for i in range(n_of_elements_plan):
            names_elements_entry.append(AutocompleteEntry(autocompleteList, root, listboxLength=8, width=50, matchesFunction=matches))            
        elem_Label.place_forget()
        elem_entry.place_forget()
        counter_of_sub = 0
        Get_names_elements_plan()
        if n_of_elements_plan != 1:
            next_button_21.place(x = 640, y = 315)
        else:
            next_button_22.place(x = 640, y = 315)

    except:
        mb.showerror("Ошибка", "Должно быть введено натуральное число")

def Next_twentyone():
    global counter_of_sub, n_of_elements_plan,names_elements_plan, names_elements_entry, name_label

    if counter_of_sub < n_of_elements_plan:
       if names_elements_entry[counter_of_sub].get() != '':
                names_elements_plan[counter_of_sub] = names_elements_entry[counter_of_sub].get()
                names_elements_entry[counter_of_sub].destroy()
                name_label.destroy()
                counter_of_sub += 1
                Get_names_elements_plan()
    if counter_of_sub == n_of_elements_plan-1:
        next_button_21.place_forget()
        next_button_22.place(x = 640, y = 315)            

def Next_twentytwo():
    global n_of_factory, names_elements_entry,name_label, counter_of_sub
    name_label.destroy() 
    names_elements_plan[counter_of_sub] = names_elements_entry[counter_of_sub].get()
    names_elements_entry[counter_of_sub].destroy()
    counter_of_sub = 0
    Quit()

#App
root = tk.Tk()

root.iconbitmap('icon.ico')
root.title("Оценка достаточности и эффективности планируемых мероприятий")
w = root.winfo_screenwidth() 
h = root.winfo_screenheight() 
w = w//2 
h = h//2 
w = w - 475
h = h - 300
root.geometry('950x600+{}+{}'.format(w, h))
root.resizable(False, False)

load = Image.open("images\\gen.jpg")
filename = ImageTk.PhotoImage(load)
background_label = tk.Label(image=filename)
background_label.place(x=0, y=0, relwidth=1, relheight=1)
filename2 = tk.PhotoImage(file = "images\\add1.gif")
tk.Label(image = filename2).place(x = 0, y= 0)
filename3 = tk.PhotoImage(file = "images\\add2.gif")
tk.Label(image = filename3, height = 76).place(x=845, y = 0)
tk.Label(text = 'ФБУН «Федеральный научный центр\nмедико-профилактических технологий управления рисками здоровью населения»', font = 'Times 12 bold',relief=tk.FLAT).place(x = 178, y = 5)

load = Image.open("images\\button.jpg")
filename4 = ImageTk.PhotoImage(load)
tk.Button(image = filename4, height = 40, width = 40,relief=tk.RAISED, command = Reference).place(x = 900, y = 550)

#Begin window
begin_button = tk.Button(height = 1, width = 15,text = 'Начать', command = Begin, font = 'Times 12 bold',relief = tk.FLAT,overrelief = tk.GROOVE,)
begin_button.place(x = 415,y = 560)
to_begin_button = tk.Button(height = 1, width = 15,text = 'В начало', command = To_begin, font = 'Times 12 bold',relief = tk.FLAT,overrelief = tk.GROOVE)
quit_button = tk.Button(height = 1, width = 15, text = 'Выйти', command = Quit, font = 'Times 12 bold',relief = tk.FLAT,overrelief = tk.GROOVE,)
quit_button.place(x=15,y = 560)

begin_frame = tk.Frame(height = 400, width = 950, background = '#87CEEB')
begin_frame.place(x = 0,y =  100)

pixelVirtual = tk.PhotoImage(width=1, height=1)

examples_frame = tk.Frame(begin_frame, height = 362, width = 250, background = '#87CEBB')
examples_frame.place(x = 690,y = 20)

load = Image.open("images\\example_image.jpg")
example_0_image = ImageTk.PhotoImage(load)
example_0_label = tk.Label(begin_frame,image = example_0_image)
example_0_label.place(x = 5  ,y = 20)

load = Image.open("images\\concentration_ex.jpg")
example_1_image = ImageTk.PhotoImage(load)
example_1_label = tk.Label(begin_frame,image = example_1_image)
example_1 = tk.Button(examples_frame, text = 'IDN, вес, возраст, с/г концентрация\nв зоне под экспозицией',font = 'Times 11',height = 45, width = 232,image=pixelVirtual,compound="c", command = exmpl_1,relief = tk.FLAT,overrelief = tk.GROOVE,activebackground = "#4682B4", background = '#E0FFFF')
example_1.place(x = 5, y = 5)

example_2_image = tk.PhotoImage(file = "images\\concentration_cl.gif")
example_2_label = tk.Label(begin_frame,image = example_2_image)
example_2 = tk.Button(examples_frame,text = 'IDN, вес, возраст, с/г концентрация\nв зоне вне экспозиции',font = 'Times 11', height = 45, width = 232,image=pixelVirtual,compound="c", command = exmpl_2,relief = tk.FLAT,overrelief = tk.GROOVE,activebackground = "#4682B4", background = '#E0FFFF')
example_2.place(x = 5, y = 65)

example_3_image = tk.PhotoImage(file = "images\\IDN_ex.gif")
example_3_label = tk.Label(begin_frame,image = example_3_image)
example_3 = tk.Button(examples_frame, text = 'IDN - диагноз\nв зоне под экспозицией',font = 'Times 12',height = 45, width = 232,image=pixelVirtual,compound="c", command = exmpl_3,relief = tk.FLAT,overrelief = tk.GROOVE,activebackground = "#4682B4", background = '#E0FFFF')
example_3.place(x = 5, y = 125)

example_4_image = tk.PhotoImage(file = "images\\IDN_cl.gif")
example_4_label = tk.Label(begin_frame,image = example_4_image)
example_4 = tk.Button(examples_frame, text = 'IDN - диагноз\nв зоне вне экспозиции',font = 'Times 12',height = 45, width = 232,image=pixelVirtual,compound="c", command = exmpl_4,relief = tk.FLAT,overrelief = tk.GROOVE,activebackground = "#4682B4", background = '#E0FFFF')
example_4.place(x = 5, y = 185)

example_5_image = tk.PhotoImage(file = "images\\Factories.gif")
example_5_label = tk.Label(begin_frame,image = example_5_image)
example_5 = tk.Button(examples_frame,text = 'Плановое снижение\nвыброса предприятиями',font = 'Times 12', height = 45, width = 232,image=pixelVirtual,compound="c", command = exmpl_5,relief = tk.FLAT,overrelief = tk.GROOVE,activebackground = "#4682B4", background = '#E0FFFF')
example_5.place(x = 5, y = 245)

example_6_image = tk.PhotoImage(file = "images\\vklad.gif")
example_6_label = tk.Label(begin_frame,image = example_6_image)
example_6 = tk.Button(examples_frame,text = 'Вклад предприятий\n в выброс веществ',font = 'Times 12', height = 45, width = 232,image=pixelVirtual,compound="c", command = exmpl_6,relief = tk.FLAT,overrelief = tk.GROOVE,activebackground = "#4682B4", background = '#E0FFFF')
example_6.place(x = 5, y = 305)

#Active window
gen_frame = tk.Frame (height = 400, width = 750)
load = Image.open("images\\blue_sky.jpg")
filename1 = ImageTk.PhotoImage(load)
load = Image.open("images\\factory.jpg")
filename5 = ImageTk.PhotoImage(load)
tk.Label(gen_frame, image = filename1, height = 400, width = 750).place(x = 0,y = 0)
tk.Label(gen_frame, text = 'Оценка достаточности и эффективности планируемых мероприятий по снижению выбросов \nзагрязняющих веществ в атмосферный воздух для митигации рисков и вреда здоровью населения\nПрограмма для ЭВМ', font = 'Times 12 bold', bg = 'SkyBlue2',relief=tk.FLAT).place(x=16,y=18)
tk.Label(gen_frame,image = filename5, height = 200, width = 750).place(x = -2, y = 100)
file_text =tk.Label(gen_frame,text = 'Выберите файл для обработки данных:', font = 'Times 14 bold',relief=tk.FLAT)
but_file = tk.Button(gen_frame, text = 'Обзор...', font = 'Times 11', command = Openfile, height = 1,relief=tk.FLAT,overrelief = tk.GROOVE)
file_label = tk.Label(gen_frame, text = '', height = 1, width = 50)
load = Image.open("images\\next.jpg")
filename6 = ImageTk.PhotoImage(load)

inf_label_1 = tk.Label(gen_frame, text = 'Язык программирования: Python\nОбъем: 333 Мб\nСкорость обработки: 5.5 Мб/мин', font = 'Times 10', bg = 'SkyBlue2',relief=tk.FLAT,justify = tk.LEFT)
inf_label_2 = tk.Label(gen_frame, text = 'Формат входных данных: .xls , .xlsx\nФормат выходных данных: .xls , .xlsx\nМаксимальный объем данных: не ограничен', font = 'Times 10',justify = tk.LEFT, bg = 'SkyBlue2',relief=tk.FLAT)
inf_label_3 = tk.Label(gen_frame, text = 'Разработчики: \nСавочкина А.А. - ФБУН ФНЦ МПТ УРЗН\nСавочкин А.И. - НИУ ВШЭ (Москва)', font = 'Times 10',justify = tk.LEFT,relief=tk.FLAT, bg = 'SkyBlue2')

next_button_0 = tk.Button(gen_frame,image = filename6, height = 66, width = 100, command = Next_null, state = tk.ACTIVE)

next_button_1 = tk.Button(gen_frame, image = filename6, height = 66, width = 100, command = Next_one, state = tk.ACTIVE)
# next_button_1.bind('<Return>', Next_one)
# next_button_1.focus_set()
next_button_2 = tk.Button(gen_frame, image = filename6, height = 66, width = 100, command = Next_two)

elem_Label = tk.Label(gen_frame, text = 'Введите количество веществ', font = 'Times 14' )
factory_Label = tk.Label(gen_frame, text = 'Введите количество предприятий', font = 'Times 14')
elem_entry = tk.Entry(gen_frame, width = 10)
factory_entry = tk.Entry(gen_frame, width = 10)


next_button_5 = tk.Button(gen_frame, image = filename6, height = 66, width = 100, command = Next_five)

next_button_10 = tk.Button(gen_frame, image = filename6, height = 66, width = 100, command = Next_ten)

next_button_11 = tk.Button(gen_frame, image = filename6, height = 66, width = 100, command = Next_eleven)

next_button_13 = tk.Button(gen_frame, image = filename6, height = 66, width = 100, command = Next_thirteen)

next_button_14 = tk.Button(gen_frame, image = filename6, height = 66, width = 100, command = Next_fourteen)

next_button_16 = tk.Button(gen_frame, image = filename6, height = 66, width = 100, command = Next_sixteen)

next_button_17 = tk.Button(gen_frame, image = filename6, height = 66, width = 100, command = Next_seventeen)

next_button_18 = tk.Button(gen_frame, image = filename6, height = 66, width = 100, command = Next_eightteen)

next_button_19 = tk.Button(gen_frame, image = filename6, height = 66, width = 100, command = Next_nineteen)

next_button_20 = tk.Button(gen_frame, image = filename6, height = 66, width = 100, command = Next_twenty)

next_button_21 = tk.Button(gen_frame, image = filename6, height = 66, width = 100, command = Next_twentyone)

next_button_22 = tk.Button(gen_frame, image = filename6, height = 66, width = 100, command = Next_twentytwo)

L_entry = tk.Entry(gen_frame, width = 10)
L_label = tk.Label(gen_frame, font = 'Times 12')
decrease_entry = tk.Entry(gen_frame, width = 10)
decrease_label = tk.Label(gen_frame, font = 'Times 12')
r_entry = tk.Entry(gen_frame, width = 10)
r_label = tk.Label(gen_frame, font = 'Times 12', text = 'Роза ветров')
v_entry = tk.Entry(gen_frame, width = 10)
v_label = tk.Label(gen_frame, font = 'Times 12', text= 'Среднегодовая скорость, км/ч')
rank_entry = tk.Entry(gen_frame, width = 10)
rank_label = tk.Label(gen_frame, font = 'Times 12', text = '')
sup_blowout_entry = tk.Entry(gen_frame, width = 10)
sup_blowout_label = tk.Label(gen_frame, font = 'Times 12', text = 'Введите коэффициент, определяющий условия горизонтального и вертикального рассеивания \nзагрязняющих веществ в атмосферном воздухе')

root.mainloop()

def c_to_doza(conc,wt,age):#count doza
    if age < 18:
        return ((conc*8*1.4)+(conc*16*0.63))*350*6/(wt*6*365)
    else:
        return ((conc*8*1.4)+(conc*16*0.63))*350*30/(wt*30*365)

def doza_to_c(doza):
    return doza*35*18*365/(350*30*(8*1.4+16*0.63))

def efficienty(delta):
    if delta <= 20:
        return 'неприемлемая'
    elif delta <= 40:
        return 'низкая'
    elif delta <= 60:
        return 'приемлемая'
    elif delta <= 80:
        return 'достаточная'
    else:
        return 'высокая'

def rounder(PDK,doza):

    if (doza <= 0.2*PDK) or (PDK == 0):
        return 1
    elif doza <= 0.3*PDK:
        if (1/(0.1*PDK)*(0.3*PDK-doza))>(1-1/(0.13*PDK)*(0.3*PDK-doza)):
            return 1
        else:    
            return 2
    elif doza <= 0.47*PDK:
        return 2
    elif doza <= 0.6*PDK:
        if (1/(0.13*PDK)*(0.6*PDK-doza))>(1-1/(0.27*PDK)*(0.6*PDK-doza)):
            return 2
        else:
            return 3
    elif doza <= 0.93*PDK:
        return 3
    elif doza <= 1.2*PDK:
        if (1/(0.27*PDK)*(1.2*PDK-doza))>(1-1/(1.73*PDK)*(2.53*PDK-doza)):
            return 3
        else:
            return 4
    elif doza <= 4.27*PDK:
        return 4
    elif doza <= 6*PDK:
        if (1/(1.73*PDK)*(2.53*PDK-doza))>(1-1/(2*PDK)*(6*PDK-doza)):
            return 4
        else:
            return 5
    else:
        return 5

def prinadlejnost(PDK,doza,bar1,bar2,bar3,bar4):

    if (doza<bar1) and (PDK != 0):
        return 0

    elif (doza<=bar2) and (PDK != 0):
        return (bar1-doza)/(bar1-bar2)

    elif (doza<=bar3) and (PDK != 0):
        return 1

    elif (doza<=bar4) and (PDK != 0):
        return (doza-bar4)/(bar3-bar4)

    else:
        return 0

def prinadlejnost_sup(PDK,doza,k):

    bar1 = 0.25*PDK
    bar2 = 0.5*PDK
    bar3 = 1*PDK
    bar4 = 5*PDK

    if k == 1:
        if (doza < 1.2*bar1) or (doza == 0) and (PDK != 0):
            if doza>=0.8*bar1:
                return abs(doza-0.8*bar1)/abs(1.2*bar1-0.8*bar1)
            else:
                return 1
        else:
            return 0

    elif k == 2:
        if (doza>=0.8*bar1) and (doza<1.2*bar2) and (PDK != 0):
            return prinadlejnost(bar2, doza, 0.2*PDK, 0.3*PDK, 0.4*PDK, 0.6*PDK)
        else:
            return 0

    elif k == 3:
        if (doza>=0.8*bar2) and (doza<1.2*bar3) and (PDK != 0):
            return prinadlejnost(bar3, doza, 0.4*PDK, 0.67*PDK, 0.94*PDK, 1.2*PDK)
        else:
            return 0

    elif k == 4:
        if (doza>=0.8*bar3) and (doza<1.2*bar4) and (PDK != 0):
            return prinadlejnost(bar4, doza, 0.8*PDK, 2.53*PDK, 4.26*PDK, 6*PDK)
        else:
            return 0

    elif k == 5:
        if (doza>=0.8*bar4) and (PDK != 0):
            if doza<1.2*bar4:
                return (doza-1.2*bar4)/(0.8*bar4-1.2*bar4)
            else:
                return 1
        else:
            return 0

def ozenka(g):

    if (g<=0.15):
        return 1

    elif g<0.25:
        if (10*(0.25-g))>(1-10*(0.25-g)):
            return 1

        else:
            return 2

    elif g<=0.35:
        return 2

    elif g<=0.45:
        if (10*(0.45-g))>(1-10*(0.45-g)):
            return 2

        else:
            return 3

    elif g<=0.55:
        return 3

    elif g<=0.65:
        if (10*(0.65-g))>(1-10*(0.65-g)):
            return 3
        else:
            return 4

    elif g<=0.75:
        return 4

    elif g<=0.85:

        if (10*(0.85-g))>(1-10*(0.85-g)):
            return 4

        else:
            return 5               

    else:
        return 5  

def convert_noz_to_dis(noz):
    for nozologia in nozologii_to_disease:
        if (noz[0:2] >= nozologia[0:2] and noz[0:2] <= nozologia[4:6]) :
            return nozologii_to_disease[nozologia]

#names_disease = ["Новообразования",
#                 "Болезни крови, кроветворных органов и отдельные нарушения, вовлекающие иммунный механизм",
#                 "Болезни эндокринной системы, расстройства питания и нарушения обмена веществ",
#                 "Болезни нервной системы",
#                 "Болезни системы кровообращения",
#                 "Болезни органов дыхания",
#                 "Болезни органов пищеварения",
#                 "Болезни кожи и подкожной клетчатки",
#                 "Болезни костно-мышечной системы и соединительной ткани",
#                 "Болезни мочеполовой системы"]

#n_of_disease = len(names_disease)

start_time = time.time()
print(n_of_elements)
#Open first sheet with data: concentration, idn, weight n age
exposition_c = pd.read_excel(excel_file, sheet_name = 0)

number_rows_ex = len(exposition_c.index)
number_columns_ex = len(exposition_c.columns)
print("reading is done")
#Open second sheet with data: concentration, idn. weight n age
clean_c = pd.read_excel(excel_file, sheet_name = 1, skiprows = 2, header = None)
number_rows_cl = len(clean_c.index)
number_columns_cl = len(clean_c.columns)

result_file_name = excel_file[:excel_file.rfind('/')] + '/' + 'results.xlsx'
writer = pd.ExcelWriter(result_file_name, engine='xlsxwriter')

#diseases_ranks_dictionary = {"Новообразования" : 10,
#                             "Болезни крови, кроветворных органов и отдельные нарушения, вовлекающие иммунный механизм" : 1,
#                             "Болезни эндокринной системы, расстройства питания и нарушения обмена веществ" : 6,
#                             "Болезни нервной системы" : 7,
#                             "Болезни системы кровообращения" : 8,
#                             "Болезни органов дыхания" : 3,
#                             "Болезни органов пищеварения" : 2,
#                             "Болезни кожи и подкожной клетчатки" : 5,
#                             "Болезни костно-мышечной системы и соединительной ткани" : 4,
#                             "Болезни мочеполовой системы" : 9}


names_disease = []
diseases_ranks_dictionary = {}

diseases_with_rank = pd.read_excel(excel_file, sheet_name=7, skiprows=1, header=None)
number_rows_diseases = len(diseases_with_rank.index)
for i in range(number_rows_diseases):
    names_disease.append(diseases_with_rank.iloc[i, 0])
    diseases_ranks_dictionary[diseases_with_rank.iloc[i, 0]] = float(diseases_with_rank.iloc[i, 2])
n_of_disease = len(names_disease)




H = [0] * n_of_factory
factories_info = pd.read_excel(excel_file, sheet_name = 6)
H = factories_info[factories_info.columns.tolist()[2]].mean()
print(H)
L = factories_info[factories_info.columns.tolist()[1]].to_numpy().tolist()
print(L)
decrease = factories_info[factories_info.columns.tolist()[3]].to_numpy().tolist()

#Table 1
print("table 1")
row_1_ex = 2
row_1_cl = number_rows_ex+3
'''
exposition_c.to_excel(writer, index=False, sheet_name='report', startrow = row_1_ex)
clean_c.to_excel(writer, index = False, sheet_name = 'report', startrow = row_1_cl)
'''
array = ["Результаты"]
table_6 = pd.DataFrame(data = array)
table_6.to_excel(writer, index=False, sheet_name='report', startrow = 0)
workbook = writer.book
worksheet = writer.sheets['report']
pd.set_option('display.max_colwidth', None)

wrap_format = workbook.add_format({
    'text_wrap': True
     })

# Reading IDN-mkb
exposition_disease = pd.read_excel(excel_file, sheet_name = 2)
IDN_mkb_ex = {}
IDN_disease_ex = {}
ex_IDN_num = 0
for i in range(len(exposition_disease.index)):
    try:
        IDN_mkb_ex[int(exposition_disease.iloc[i,0])] = exposition_disease.iloc[i,1]
        #IDN_disease_ex[int(exposition_disease.iloc[i,0])] = convert_noz_to_dis(exposition_disease.iloc[i,1])
        IDN_disease_ex[int(exposition_disease.iloc[i,0])] = exposition_disease.iloc[i,1].strip()
    except:
        1
        #print('Количество людей IDN - болезнь не соответствует количеству людей IDN - концентрации (в зоне под экспозицией)')

#exposition_disease.to_excel(writer, index = False, header = False, sheet_name = 'report', startrow = row_3_ex)
clean_disease = pd.read_excel(excel_file, sheet_name = 3, header = None)
IDN_disease_cl = {}
IDN_mkb_cl = {}
for i in range(len(clean_disease) - 1):
    IDN_mkb_cl[int(clean_disease.iloc[i + 1,0])] = clean_disease.iloc[i + 1,1]
    #IDN_disease_cl[int(clean_disease.iloc[i + 1,0])] = convert_noz_to_dis(clean_disease.iloc[i + 1,1])
    IDN_disease_cl[int(clean_disease.iloc[i + 1,0])] = clean_disease.iloc[i + 1,1].strip()
    #print('Количество людей IDN - болезнь не соответствует количеству людей IDN - концентрации (вне зоны экспозиции)')
print(IDN_disease_ex)
print("second reading is done")

diseases = []
for i in range(len(names_disease)):
    if names_disease[i] in list(IDN_disease_ex.values()) or names_disease[i]  in list(IDN_disease_cl.values()):
        diseases.append(names_disease[i])
names_disease = diseases
n_of_disease = len(names_disease)

#Reading data from 'bad' place
IDN_ex = []
wt_ex = []
age_ex = []
c_ex = [[0] * n_of_elements for i in range(len(IDN_mkb_ex))]
av_c_ex = [0] * n_of_elements

ADDch_ex = [[0] * n_of_elements for i in range(len(IDN_mkb_ex))]

for i in range(n_of_elements):
    names_elements.append(exposition_c.iloc[0][i+3])

now = datetime.datetime.now()
index = 0
for i in range(number_rows_ex-2):
   if int(exposition_c.iloc[i+2][0]) not in IDN_mkb_ex:
        continue
   IDN_ex.append(int(exposition_c.iloc[i+2][0]))
   wt_ex.append(float(exposition_c.iloc[i+2][1]))
   age_ex.append(int(exposition_c.iloc[i+2][2]))
   for j in range(n_of_elements):
       c_ex[index][j] = float(exposition_c.iloc[i+2][j+3])
       av_c_ex[j] += c_ex[index][j]
       ADDch_ex[index][j] = c_to_doza(c_ex[index][j],wt_ex[index],age_ex[index])
   index += 1

for i in range(n_of_elements):
    av_c_ex[i] /= len(IDN_mkb_ex) 
#Reading data from 'good' place
IDN_cl = []
wt_cl = []
age_cl = []
c_cl = [[0] * n_of_elements for i in range(len(IDN_mkb_cl))]
ADDch_cl = [[0] * n_of_elements for i in range(len(IDN_mkb_cl))]
index = 0
for i in range(number_rows_cl - 1):
   if int(clean_c.iloc[i][0]) not in IDN_mkb_cl:
        continue
   IDN_cl.append(int(clean_c.iloc[i][0]))
   wt_cl.append(float(clean_c.iloc[i][1]))
   age_cl.append(int(clean_c.iloc[i][2]))
   for j in range(n_of_elements):
       c_cl[index][j] = float(clean_c.iloc[i][j+3])
       ADDch_cl[index][j] = c_to_doza(c_cl[index][j],wt_cl[index],age_cl[index])
   index += 1

#Table 2
print("table 2")
row_2_ex = number_rows_ex+number_rows_cl+5+row_1_ex
row_2_cl = number_rows_ex+number_rows_cl+6+row_1_cl
'''
exposition_d = exposition_c.drop(['Вес, кг', 'Возраст, лет'], axis = 1)

for i in range(number_rows_ex-1):
   for j in range(n_of_elements):
       exposition_d.iloc[i+1,j+1] = ADDch_ex[i][j]

#exposition_d.to_excel(writer, index = False, sheet_name = 'report', startrow = row_2_ex, startcol = 0)


clean_d = clean_c.drop([1,2], axis = 1)

add_column = []
for i in range(n_of_elements+1):
    add_column.append(i)
clean_d.columns = add_column

for i in range(number_rows_cl):
   for j in range(n_of_elements):
       clean_d.iloc[i,j+1] = ADDch_cl[i][j]
#clean_d.to_excel(writer, index = False, header = False,sheet_name = 'report', startrow = row_2_cl)
'''
#Table 3
print("table 3")
row_3_ex = number_rows_ex+number_rows_cl+5+row_2_ex
row_3_cl = number_rows_ex+number_rows_cl+3+row_2_cl 

#clean_disease.to_excel(writer, index = False,header = False, sheet_name = 'report', startrow = row_3_cl)

#Table 4
print("table 4")
row_4_ex = number_rows_ex+number_rows_cl+5+row_3_ex
row_4_cl = number_rows_ex+number_rows_cl+6+row_3_cl 
'''
exposition_IDN_d = exposition_d

exposition_IDN_d[n_of_elements+1] = ""
for i in range(number_rows_ex-1):
    try:
        exposition_IDN_d.iloc[i+1,n_of_elements+1] = IDN_disease_ex[exposition_IDN_d.iloc[i+1,0]]
    except:
        1
        #print('Количество людей IDN - болезнь не соответствует количеству людей IDN - концентрации (в зоне экспозиции)')
      
exposition_IDN_d.to_excel(writer, index = False, header = False, sheet_name = 'report', startrow = row_4_ex)

clean_IDN_d = clean_d
clean_IDN_d[n_of_elements+1] = ""

for i in range(number_rows_cl):
    try:
        clean_IDN_d.iloc[i,n_of_elements+1] = IDN_disease_cl[clean_IDN_d.iloc[i,0]]
    except:
        1
        #print('Количество людей IDN - болезнь не соответствует количеству людей IDN - концентрации (вне зоны экспозиции)')
               
#clean_IDN_d.to_excel(writer, index = False, header = False, sheet_name = 'report', startrow = row_4_cl)
'''
#Table 5
print("table 5")

#Add list of disease as number

#row_5_ex = number_rows_ex+number_rows_cl+5+row_4_ex
row_5_ex = 5
#row_5_cl = number_rows_ex+number_rows_cl+5+row_4_ex+n_of_disease+3  
row_5_cl = row_5_ex + n_of_disease + 3

ex_disease_d = [[0] * n_of_elements for i in range(n_of_disease)]
temp_d = [0] * n_of_elements
q = 0
count_disease_ex = [0] * n_of_disease #number of people 
disease_ex_gl = [0] * n_of_elements # general value of doza
count_ex = 0
count_ex_d = 0 # number non-zero elements for average value

for i in range(n_of_disease):
    for j in IDN_disease_ex:
        for m in range(n_of_elements):
            temp_d[m] = ADDch_ex[q][m]
            if IDN_disease_ex[j] == names_disease[i]:
                ex_disease_d[i][m] += temp_d[m]
                count_disease_ex[i] += 1
        q += 1
    count_disease_ex[i] /= n_of_elements
    q = 0
    count_ex += count_disease_ex[i]
    if count_disease_ex[i] != 0:
        count_ex_d += 1


for i in range(n_of_disease):
    for j in range(n_of_elements):
        try:
            ex_disease_d[i][j] /= count_disease_ex[i]
            disease_ex_gl[j] += ex_disease_d[i][j]
        except:
            None

ex_disease = pd.DataFrame(data = ex_disease_d)
index_dis = []
for i in range(n_of_disease):
    index_dis.append(names_disease[i])

ex_disease.columns = names_elements
ex_disease.index = index_dis
ex_disease.loc['По совокупности классов болезней'] = None

aver_ex_d = [0] * n_of_elements

for i in range(n_of_elements):
    if count_ex_d == 0:
        ex_disease.iloc[n_of_disease,i] = 0
        aver_ex_d[i] = 0
    else:
        ex_disease.iloc[n_of_disease,i] = disease_ex_gl[i] / count_ex_d
        aver_ex_d[i] = disease_ex_gl[i] / count_ex_d
    
for i in range(n_of_elements):
    disease_ex_gl[i] /= count_ex


ex_disease.to_excel(writer, sheet_name = 'report', startrow = row_5_ex, startcol = 2)


cl_disease_d = [[0] * n_of_elements for i in range(n_of_disease)]
temp_d = [0] * n_of_elements
q = 0
count_disease_cl = [0] * n_of_disease
count_cl = 0
disease_cl_gl = [0] * n_of_elements
count_cl_d = 0
for i in range(n_of_disease):
    for j in IDN_disease_cl:
        for m in range(n_of_elements):
            temp_d[m] = ADDch_cl[q][m]
            if IDN_disease_cl[j] == names_disease[i]:
                cl_disease_d[i][m] += temp_d[m]
                count_disease_cl[i] += 1
        q += 1
    count_disease_cl[i] /= n_of_elements
    q = 0
    count_cl += count_disease_cl[i]
    if count_disease_cl[i] != 0:
        count_cl_d += 1
for i in range(n_of_disease):
    for j in range(n_of_elements):
        try:
            cl_disease_d[i][j] /= count_disease_cl[i]
            disease_cl_gl[j] += cl_disease_d[i][j]
        except:
            None
            
cl_disease = pd.DataFrame(data = cl_disease_d)

cl_disease.index = index_dis
cl_disease.loc['По совокупности классов болезней'] = None

aver_cl_d = [0] * n_of_elements

for i in range(n_of_elements):
    cl_disease.iloc[n_of_disease,i] = disease_cl_gl[i] / count_cl_d 
    aver_cl_d[i] = disease_cl_gl[i] / count_cl_d  
    
for i in range(n_of_elements):
    disease_cl_gl[i] /= count_cl

cl_disease.to_excel(writer, sheet_name = 'report', header = False,startrow = row_5_cl, startcol = 2)

print("table 6")
#Table 6
#row_6 = row_5_cl + n_of_disease+ 4
#row_6 = 4
row_6 = row_5_cl + n_of_disease + 4

PDK_elements = [0]*n_of_elements

for i in range(n_of_elements):
    try:
        PDK_elements[i] = substance_to_PDK[names_elements[i]]
        print(PDK_elements[j])
               
    except:
        print(names_elements[i])

ex_round = [[0] * n_of_elements for i in range(n_of_disease+1)]
for i in range(n_of_disease):
    for j in range(n_of_elements):

            ex_round[i][j] = rounder(PDK_elements[j],ex_disease_d[i][j])
for j in range(n_of_elements):
    ex_round[i+1][j] = rounder(PDK_elements[j],aver_ex_d[j])


table_6 = pd.DataFrame(data = ex_round)

table_6.index = index_dis+['По совокупности болезней']
table_6.columns = names_elements
table_6.to_excel(writer, sheet_name = 'report', startrow = row_6, startcol = 2)

#Table 7
row_7 = row_6 + n_of_disease + 5

ranking = [0]*n_of_disease

# Для нозологий

#for i in range(n_of_disease):
#    if names_disease[i] not in diseases_ranks_dictionary.keys():
#        print(names_disease[i])

#for i in range(n_of_disease):
#    ranking[i] = diseases_ranks_dictionary[names_disease[i]]



factor = [0]*n_of_disease

for i in range(n_of_disease):
    factor[i] =  2*(n_of_disease-ranking[i]+1)/(n_of_disease*(n_of_disease+1))

r_sup = [0.125,0.3,0.5,0.7,0.875]

p_general = [[0]*5 for i in range(n_of_disease)]

for k in range(5):
    for i in range(n_of_disease):
        for j in range(n_of_elements):
            p_general[i][k] += factor[i]*abs(prinadlejnost_sup(PDK_elements[j],ex_disease_d[i][j],k+1))

R = [0]*n_of_disease 
global_R = 0
for i in range(n_of_disease):
    for k in range(5):
       R[i] += p_general[i][k]*r_sup[k]
    global_R += R[i]
global_R /= n_of_disease
p_g_concrete = [[0]*n_of_elements for i in range(n_of_disease)]

for i in range(n_of_disease):
    for j in range(n_of_elements):
        for k in range(5):
            p_g_concrete[i][j] += r_sup[k]*factor[i]*prinadlejnost_sup(PDK_elements[j],ex_disease_d[i][j],k+1) 
        if R[i] == 0:
            p_g_concrete[i][j] = 0
        else:
            p_g_concrete[i][j] = round(p_g_concrete[i][j]/R[i]*100,2)
table_7 = pd.DataFrame(data = p_g_concrete)
table_7.index = index_dis
table_7.columns = names_elements
table_7.to_excel(writer, sheet_name = 'report', startrow = row_7, startcol = 2)

#Table 8
row_8 = row_7 + n_of_disease + 3
table_8 = pd.DataFrame(data = R)
table_8.index = index_dis
table_8.to_excel(writer, sheet_name = 'report', startrow = row_8, startcol = 2)

#Table 9

row_9 = row_8 + n_of_disease + 3
decrease_elements = []
additional = [0,1,2,3,4,5]
for i in range(n_of_elements):
    if rounder(PDK_elements[i],aver_ex_d[i]) >= 2:
        decrease_elements.append(names_elements[i])
table_9 = pd.DataFrame(columns = additional)
for i in range(len(decrease_elements)):
    table_9.loc[i,2] = decrease_elements[i]
    table_9.loc[i,4] = decrease_elements[i]
for i in range(len(names_elements_plan)):
    table_9.loc[i,0] = names_elements_plan[i]
if len(decrease_elements) > len(names_elements_plan):
    for i in range(len(names_elements_plan),len(decrease_elements)):
        table_9.loc[i,0] = ''
else:
    for i in range(len(decrease_elements),len(names_elements_plan)):
        table_9.loc[i,2] = ''
        table_9.loc[i,4] = ''       

table_9.to_excel(writer, sheet_name = 'report', startrow = row_9, index = False)

#Table 10
if len(decrease_elements) > len(names_elements_plan):
    row_10 = row_9 + len(decrease_elements) + 3
else:
    row_10 = row_9 + len(names_elements_plan) + 3

blowout = [0] * n_of_elements

blowout_sum = 0
determined_factor_A_F_m_n_nu = A*F*m*n*nu
for i in range(n_of_elements):
    blowout[i] = av_c_ex[i]*H*H*1*31.536/determined_factor_A_F_m_n_nu
    blowout_sum += blowout[i]
table_10 = pd.DataFrame(columns = additional)
for i in range(n_of_elements):
    table_10.loc[i,0] = names_elements[i]
    table_10.loc[i,3] = blowout[i]
table_10.loc[i+1,0] = 'Совокупно по всем веществам'    
table_10.loc[i+1,3] = blowout_sum
table_10.to_excel(writer, sheet_name = 'report', startrow = row_10, index = False)

#Table 11
row_11 = row_10 + n_of_elements + 6

factories = pd.read_excel(excel_file, sheet_name = 4, skiprows = 2, header = None)

contribution = pd.read_excel(excel_file, sheet_name = 5, skiprows = 1, header = None)
vklad = [[0]*n_of_elements for i in range(n_of_factory)]
for i in range(n_of_factory):
    for j in range(n_of_elements):
        vklad[i][j] = 1#float(contribution.iloc[i,j+1])/100

blowout_factory_plan =[[0]*n_of_factory for i in range(n_of_elements)]
for i in range(n_of_elements):
    for j in range(n_of_factory):
        try:
            blowout_factory_plan[i][j] = float(factories.loc[i,j+1])
        except:
            blowout_factory_plan[i][j] = '-'
additional_factory = []
for i in range(3*n_of_factory+1):
    additional_factory.append(i)
    
table_11 = pd.DataFrame(columns = additional_factory)

print(blowout_factory_plan)

blowout_factory_real = [[0]*n_of_factory for i in range(n_of_elements)]
for i in range(n_of_elements):
    for j in range(n_of_factory):
        blowout_factory_real[i][j] = blowout[i]*(1/L[j])*r*V*(16*350)*vklad[j][i] 

for i in range(n_of_elements):
    for j in range(0,n_of_factory*3-2,3):
        table_11.loc[i,j+1] = blowout_factory_plan[i][j//3]
        table_11.loc[i,j+2] = blowout_factory_real[i][j//3]
        if (blowout_factory_plan[i][j//3] != '-' and blowout_factory_plan[i][j//3] != 0):
            table_11.loc[i,j+3] = round((-blowout_factory_real[i][j//3]+blowout_factory_plan[i][j//3])/blowout_factory_plan[i][j//3]*100,2)
        else:
            table_11.loc[i,j+3] = '-'
for i in range(n_of_elements):    
    table_11.loc[i,0] = names_elements[i]

factory_name = [0]*n_of_factory

for i in range(n_of_factory):
    factory_name[i] = 'Предприятие ' + str(i+1)

table_11.to_excel(writer, sheet_name = 'report', startrow = row_11, index = False, startcol = 2)

#Table 12
row_12 = row_11 + n_of_elements + 4

table_12 = pd.DataFrame(columns = names_elements, index = factory_name)

for i in range(n_of_factory):
    for j in range(n_of_elements):
        if vklad[i][j] != 0:
            try:
                table_12.iloc[i,j] = doza_to_c(PDK_elements[j])*H*H*1*31.536/(A*F*m*n*nu)*(1/L[i])*r*V*(24*350)*vklad[i][j]
            except:
                table_12.iloc[i,j] = 0
        else:
            table_12.iloc[i,j] = '-'
table_12.to_excel(writer, sheet_name = 'report', startrow = row_12)

#Table 13
row_13 = row_12 + n_of_factory + 4

good_doza = [[0]*n_of_elements for i in range(n_of_factory)]

for i in range(n_of_factory):
    for j in range(n_of_elements):
        if vklad[i][j] != 0:
            good_doza[i][j] = PDK_elements[j] * vklad[i][j] * (1-decrease[i])
        else:
            good_doza[i][j] = '-'

table_13 = pd.DataFrame(columns = names_elements, index = factory_name, data = good_doza)

table_13.to_excel(writer, sheet_name = 'report', startrow = row_13)

print("table 14")

#Table 14
row_14 = row_13 + n_of_factory + 4

factory_doza = [[0]*n_of_elements for i in range(n_of_factory)]

for i in range(n_of_factory):
    for j in range(n_of_elements):
        if vklad[i][j] != 0:
            factory_doza[i][j] = aver_ex_d[j]* vklad[i][j] * (1-decrease[i])
        else:
            factory_doza[i][j] = '-'

general_factory_doza = [0]*n_of_elements
rang_general_f_doza = [0]*n_of_elements
count_of_nonull_factory = [0]*n_of_elements

for i in range(n_of_elements):
    for j in range(n_of_factory):
        if factory_doza[j][i] != '-':
            general_factory_doza[i] += factory_doza[j][i]
            count_of_nonull_factory[i] += 1
for i in range(n_of_elements):
    try:
        rang_general_f_doza[i] = rounder(PDK_elements[i],general_factory_doza[i])
    except:
        rang_general_f_doza[i] = 1
table_14 = pd.DataFrame(index = names_elements, columns = [0,1,2,3,4,5])
for i in range(n_of_elements):
    table_14.iloc[i,0] = rounder(PDK_elements[i],aver_ex_d[i])
    table_14.iloc[i,3] = rang_general_f_doza[i]

table_14.to_excel(writer, sheet_name = 'report', startrow = row_14)

#Table 15
row_15 = row_14 + n_of_elements + 4

decrease_blowout_all = [0]*n_of_elements
decrease_blowout = [[0]*n_of_elements for i in range(n_of_factory)]
for i in range(n_of_factory):
    for j in range(n_of_elements):
        try:
            if doza_to_c(0.25*PDK_elements[j])*H*H*1*31.536/(A*F*m*n*nu)*(1/L[i])*r*V*(24*350)*vklad[i][j] < blowout_factory_real[j][i]*(1-decrease[i]):
                decrease_blowout[i][j] = blowout_factory_real[j][i]*(1-decrease[i])-doza_to_c(0.25*PDK_elements[j])*H*H*1*31.536/(A*F*m*n*nu)*(1/L[i])*r*V*(24*350)*vklad[i][j]
                decrease_blowout_all[j] += decrease_blowout[i][j]
            else:
                decrease_blowout[i][j] = '-'
        except:
            decrease_blowout[i][j] = '-'

table_15 = pd.DataFrame(data = decrease_blowout, index = factory_name, columns = names_elements)
table_15.to_excel(writer, sheet_name = 'report', startrow = row_15)

#Table 16
row_16 = row_15 + n_of_factory + 3

table_16 = pd.DataFrame(columns = [0,1,2,3,4,5], index = names_disease)
for i in range(n_of_disease):
    table_16.iloc[i,0] = names_disease[i]

decrease_doza = [0]*n_of_elements

for i in range(n_of_elements):
    for j in range(n_of_factory):
        if vklad[j][i] != 0:
            decrease_doza[i] += aver_ex_d[i] * vklad[j][i] * (1-decrease[j])
    if decrease_doza[i] != 0: 
        decrease_doza[i] /= count_of_nonull_factory[i]

p_general_dec = [[0]*5 for i in range(n_of_disease)]

for k in range(5):
    for i in range(n_of_disease):
        for j in range(n_of_elements):
            p_general_dec[i][k] += factor[i]*prinadlejnost_sup(PDK_elements[j],decrease_doza[j],k+1) #Need some changes

R_dec = [0]*n_of_disease #NEED CHANGE
global_dec_R = 0
for i in range(n_of_disease):
    for k in range(5):
       R_dec[i] += p_general_dec[i][k]*r_sup[k]
    global_dec_R += R_dec[i]
global_dec_R /= n_of_disease
for i in range(n_of_disease):
    table_16.iloc[i,3] = R_dec[i]
table_16.to_excel(writer, sheet_name = 'report', startrow = row_16,index = False)

#Table 17
row_17 = row_16 + n_of_disease + 4

table_17 = pd.DataFrame(columns = [0,1,2,3,4,5,6,7], index = [0])
table_17.iloc[0,0] = round(global_R,2)
table_17.iloc[0,2] = round(global_dec_R,2)
delta_R = global_R - global_dec_R

table_17.iloc[0,4] = round(delta_R*100/global_R,2)
table_17.iloc[0,6] = efficienty(delta_R/global_R*100)

table_17.to_excel(writer, sheet_name = 'report', startrow = row_17,index = False,float_format="%.2f")

print("table 18")

#Table 18
row_18 = row_17 + 4

blowout_factory_real_sum = [0]*n_of_elements

for i in range(n_of_elements):
    for j in range(n_of_factory):
        blowout_factory_real_sum[i] += blowout_factory_real[i][j]

addit_decrease = []

for i in range(n_of_elements):
    if decrease_blowout_all[i] != '-' and blowout_factory_real_sum[i] != 0:
        addit_decrease.append(round(decrease_blowout_all[i]/blowout_factory_real_sum[i]*100,2))
aux_mas = []
for i in range(n_of_elements):
    aux_mas.append(i)

table_18 = pd.DataFrame(columns = [0,1,2,3], index = aux_mas)
for i in range(n_of_elements):
    table_18.iloc[i,0] = names_elements[i]
    #table_18.iloc[i,2] = addit_decrease[i]
table_18.to_excel(writer, sheet_name = 'report', startrow = row_18, index = False)

merge_format = workbook.add_format({
    'bold': 1,
    'border': 2,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': True
    })

main_format = workbook.add_format({
    'valign': 'vcenter',
    'align': 'center',
    'border':2,
    'text_wrap': True,
     })

names_format = workbook.add_format({
    'align': 'center',
    'font_size' : '12',
    'bold' : True,
    'italic': True,
    'text_wrap': True
    })

table_format = workbook.add_format({
    'align': 'center',
    'font_size': '12',
    'bg_color': '#CCCCFF'
    })

danger_format_1 = workbook.add_format({
    'align': 'center',
    'bg_color': '#00CCFF',
    'border': 2,
    'text_wrap': True,
    })

danger_format_2 = workbook.add_format({
    'align': 'center',
    'bg_color': '#99CC00',
    'border': 2,
    'text_wrap': True,
    })

danger_format_3 = workbook.add_format({
    'align': 'center',
    'bg_color': '#FFCC00',
    'border': 2,
    'text_wrap': True,
    })

danger_format_4 = workbook.add_format({
    'align': 'center',
    'bg_color': '#FF9900',
    'border': 2,
    'text_wrap': True,
    })

danger_format_5 = workbook.add_format({
    'align': 'center',
    'bg_color': '#ff0000',
    'border': 2,
    'text_wrap': True,
    })

worksheet.set_column(0,number_columns_ex*3, 12, main_format)
'''
#Table 1
worksheet.write(row_1_ex-2,0, 'Таблица 1',table_format)
worksheet.merge_range(row_1_ex-2,1,row_1_ex-2,number_columns_cl-1, 'Средняя годовая концентрация загрязняющего вещества в атмосферном воздухе, вес и возраст индивидуума зон под экспозицией  и вне экспозиции', names_format)
worksheet.merge_range(row_1_ex-1,0,row_1_ex,0, 'ИДН индивидуума с адресной привязкой к контрольной точке',merge_format)
worksheet.merge_range(row_1_ex-1,1,row_1_ex,1, 'Вес, кг',merge_format)   
worksheet.merge_range(row_1_ex-1,2,row_1_ex,2, 'Возраст, лет',merge_format)
for i in range(n_of_elements):
    worksheet.write(row_1_ex,i+3, names_elements[i], main_format)
worksheet.merge_range(row_1_ex+1,0,row_1_ex+1,number_columns_ex-1,'В зоне под экспозицией', merge_format)
worksheet.merge_range(row_1_ex-1,3,row_1_ex-1,number_columns_ex-1, 'Средняя годовая концентрация (Сс,г,ji), мг/м3',merge_format)
worksheet.merge_range(row_1_cl,0,row_1_cl,number_columns_cl-1, 'Вне зоны экспозиции', merge_format)
worksheet.set_row(row_1_cl, 30)
worksheet.set_row(row_1_ex+1,30)
worksheet.set_row(row_1_ex-2,60)
#Table 2
worksheet.write(row_2_ex-2,0,'Таблица 2',table_format)
worksheet.merge_range(row_2_ex-2,1,row_2_ex-2,number_columns_cl-3,'Хроническая средняя суточная доза вещества при аэрогенном поступлении в организм индивидуума зон под экспозицией и вне экспозиции', names_format)
worksheet.merge_range(row_2_ex-1,0,row_2_ex,0, 'ИДН индивидуума с адресной привязкой к контрольной точке', merge_format )
worksheet.merge_range(row_2_ex-1,1,row_2_ex-1,number_columns_cl-3, 'Хроническая  средняя суточная доза вещества (ADDch), усредненная на экспозицию с частотой воздействия 365 дней (1 год), мг/(кг*сутки)', merge_format)
for i in range(n_of_elements):
    worksheet.write(row_2_ex,i+1, names_elements[i], main_format)
worksheet.merge_range(row_2_ex+1,0,row_2_ex+1,number_columns_ex-3, 'В зоне под экспозицией', merge_format)
worksheet.merge_range(row_2_cl-1,0,row_2_cl-1,number_columns_cl-3, 'Вне зоны экспозиции', merge_format)
worksheet.set_row(row_2_ex+1, 30)
worksheet.set_row(row_2_cl-1, 30)
worksheet.set_row(row_2_ex-2,60)
#Table 3
worksheet.write(row_3_ex-2,0,'Таблица 3',table_format)
worksheet.merge_range(row_3_ex-2,1,row_3_ex-2,7, 'Основной диагноз впервые выявленного хронического заболевания по классу болезней для каждого индивидуума в каждой контрольной точке зон под экспозицией и вне  экспозиции за соответствующий период наблюдения (1 год)', names_format)
worksheet.write(row_3_ex-1,0,'ИДН индивидуума с адресной привязкой к контрольной точке', merge_format)
worksheet.write(row_3_ex-1,1,'Диагноз заболевания (код по МКБ)', merge_format)
worksheet.merge_range(row_3_ex-1,2,row_3_ex-1,7, 'Класс болезни в соответствии с диагнозом заболевания по МКБ', merge_format)
worksheet.merge_range(row_3_ex,0,row_3_ex,7, 'В зоне под экспозицией', merge_format)
for i in range(number_rows_ex-1):
    try:
        worksheet.merge_range(row_3_ex+i+1,2,row_3_ex+i+1,7,exposition_disease.iloc[i+1,2], main_format)
    except:
        # 'Новообразование (почки)'
        # 'Новообразвания (почки)'
        1
        #print('Количество людей IDN - болезнь не соответствует количеству людей IDN - концентрации (зона экспозиции)')
        break
worksheet.merge_range(row_3_cl,0,row_3_cl,7, 'Вне зоны экспозии', merge_format)
for i in range(number_rows_cl):
    try:
        worksheet.merge_range(row_3_cl+i+1,2,row_3_cl+i+1,7,clean_disease.iloc[i+1,2], main_format)
    except:
        1
        #print('Количество людей IDN - болезнь не соответствует количеству людей IDN - концентрации (вне зоны экспозиции)')
        break
worksheet.set_row(row_3_cl, 30)
worksheet.set_row(row_3_ex, 30)
worksheet.set_row(row_3_ex-2,60)
#Table 4
worksheet.write(row_4_ex-3,0,'Таблица 4',table_format)
worksheet.merge_range(row_4_ex-3,1,row_4_ex-3,n_of_elements+5,'Хроническая средняя суточная доза вещества при аэрогенном поступлении в организм  индивидуума и класс болезни, к которому отнесен  диагноз впервые выявленного хронического заболевания, для каждого индивидуума в каждой контрольной точке зон под экспозицией и вне экспозиции за соответствующий период наблюдения (1 год)', names_format)
for i in range(number_rows_ex-1):
    try:
        worksheet.merge_range(row_4_ex+i+1,n_of_elements+1,row_4_ex+i+1,n_of_elements+5,IDN_disease_ex[exposition_IDN_d.iloc[i+1,0]], main_format)
    except:
        1
        #print('Количество людей IDN - болезнь не соответствует количеству людей IDN - концентрации (зона экспозиции)')
 
for i in range(number_rows_cl):
    try:
        worksheet.merge_range(row_4_cl+i,n_of_elements+1,row_4_cl+i,n_of_elements+5,IDN_disease_cl[clean_disease.iloc[i+1,0]], main_format)
    except:
        1
        #print('Количество людей IDN - болезнь не соответствует количеству людей IDN - концентрации (вне зоны экспозиции)')

worksheet.merge_range(row_4_ex-2,0,row_4_ex-1,0,'ИДН индивидуума с адресной привязкой к контрольной точке', merge_format)
worksheet.merge_range(row_4_ex-2,1,row_4_ex-2,n_of_elements,'Хроническая  средняя суточная доза вещества (ADDch), усредненная на экспозицию с частотой воздействия 365 дней (1 год), мг/(кг*сутки)', merge_format)
worksheet.merge_range(row_4_ex-2,n_of_elements+1,row_4_ex-1,n_of_elements+5,'Класс болезни в соответствии с диагнозом заболевания (код по МКБ)', merge_format)
for i in range(n_of_elements):
    worksheet.write(row_4_ex-1,i+1,names_elements[i],merge_format)
worksheet.merge_range(row_4_ex,0,row_4_ex,n_of_elements+5,'В зоне под экспозицией', merge_format)
worksheet.merge_range(row_4_cl-1,0,row_4_cl-1, n_of_elements+5, 'Вне зоны экспозиции', merge_format)
worksheet.set_row(row_4_cl-1, 30)
worksheet.set_row(row_4_ex, 30)
worksheet.set_row(row_4_ex-3,60)
'''
#Table 5
worksheet.write(row_5_ex-3,0,'Таблица 5',table_format)
worksheet.merge_range(row_5_ex-3,1,row_5_ex-3,n_of_elements+2,'Классы болезней, к которым отнесены  впервые выявленные хронические заболевания, соотнесенные с хронической средней суточной дозой вещества, детерминирующей совокупный спектр негативных ответов по всем выявленным классам болезней за соответствующий период наблюдения (1 год)',names_format)
worksheet.merge_range(row_5_ex-2,0,row_5_ex-1,2, 'Наименование класса болезней', merge_format)
worksheet.merge_range(row_5_ex-2,3,row_5_ex-2,n_of_elements+2,'Хроническая  средняя суточная доза вещества (ADDch), усредненная на экспозицию с частотой воздействия 365 дней (1 год), мг/(кг*сутки)', merge_format)
for i in range(n_of_disease):
    worksheet.merge_range(row_5_ex+i+1,0,row_5_ex+i+1,2,names_disease[i],merge_format)
worksheet.merge_range(row_5_ex+n_of_disease+1,0,row_5_ex+n_of_disease+1,2,'По совокупности болезней',merge_format)
worksheet.merge_range(row_5_ex,0,row_5_ex,n_of_elements+2, 'В зоне под экспозицией', merge_format)

for i in range(n_of_elements):
    worksheet.write(row_5_ex-1,i+3,names_elements[i], merge_format)

for i in range(n_of_disease):
    worksheet.merge_range(row_5_cl+i,0,row_5_cl+i,2,names_disease[i],merge_format)
worksheet.merge_range(row_5_cl+n_of_disease,0,row_5_cl+n_of_disease,2,'По совокупности болезней',merge_format)

worksheet.merge_range(row_5_cl-1,0,row_5_cl-1,n_of_elements+2, 'Вне экспозиции', merge_format)
worksheet.set_row(row_5_ex,30)
worksheet.set_row(row_5_ex-3,60)

#Table 6
worksheet.write(row_6-2,0,'Таблица 7',table_format)
worksheet.merge_range(row_6-2,1,row_6-2,n_of_elements+2,'Перечень веществ, представляющих потенциальную опасность причинения вреда здоровью при аэрогенных экспозициях', names_format)
worksheet.merge_range(row_6-1,0,row_6,2,'Наименование класса болезней', merge_format)
worksheet.merge_range(row_6-1,3,row_6-1,n_of_elements+2,'Ранг степени опасности причинения вреда здоровью (Rang)', merge_format)
for i in range(n_of_disease):
    worksheet.merge_range(row_6+i+1,0,row_6+i+1,2,names_disease[i],merge_format)
for i in range(n_of_elements):
    if table_6.iloc[n_of_disease,i] == 1:
        worksheet.write(row_6,3+i, names_elements[i],danger_format_1)
    elif table_6.iloc[n_of_disease,i] == 2:
        worksheet.write(row_6,3+i, names_elements[i],danger_format_2)
    elif table_6.iloc[n_of_disease,i] == 3:
        worksheet.write(row_6,3+i, names_elements[i],danger_format_3)
    elif table_6.iloc[n_of_disease,i] == 4:
        worksheet.write(row_6,3+i, names_elements[i],danger_format_4)
    elif table_6.iloc[n_of_disease,i] == 5:
        worksheet.write(row_6,3+i, names_elements[i],danger_format_5)
  
worksheet.merge_range(row_6+n_of_disease+1,0,row_6+n_of_disease+1,2,'По совокупности болезней',merge_format)
worksheet.set_row(row_6-2,60)
#Table 7

worksheet.write(row_7-2,0,'Таблица 10',table_format)
worksheet.merge_range(row_7-2,1,row_7-2,n_of_elements+2,'Вклад вещества в причиненный вред здоровью', names_format)
worksheet.merge_range(row_7-1,0,row_7,2,'Наименование класса болезней', merge_format)
worksheet.merge_range(row_7-1,3,row_7-1,n_of_elements+2,'Вклад вещества в причиненный вред здоровью, %', merge_format)
for i in range(n_of_disease):
    worksheet.merge_range(row_7+i+1,0,row_7+i+1,2,names_disease[i],merge_format)
for i in range(n_of_elements):
    worksheet.write(row_7,3+i, names_elements[i],merge_format)
worksheet.set_row(row_7-2,60)
#Table 8
    
worksheet.write(row_8-1,0,'Таблица 11', table_format)
worksheet.merge_range(row_8-1,1,row_8-1,5,'Причиненный вред здоровью в виде негативного эффекта (по конкретному классу болезней), детерминированного аэрогенной экспозицией до проведения воздухоохранных мероприятия', names_format)
worksheet.merge_range(row_8,0,row_8,2,'Наименование класса болезней', merge_format)
worksheet.merge_range(row_8,3,row_8,5,'Величина причиненного вреда здоровью, детерминированный аэрогенной экспозицией',merge_format)
for i in range(n_of_disease):
    worksheet.merge_range(row_8+i+1,0,row_8+1+i,2,names_disease[i],merge_format)
    worksheet.merge_range(row_8+i+1,3,row_8+i+1,5,table_8.iloc[i,0],merge_format)
worksheet.set_row(row_8-1, 60)

#Table 9
worksheet.write(row_9-1,0,'Таблица 12', table_format)
worksheet.merge_range(row_9-1,1,row_9-1,5, 'Список веществ, выбросы которых подлежат  обязательному регулированию по критериям причиненного вреда здоровью',names_format)
worksheet.merge_range(row_9,0,row_9,1,'Перечень веществ,  выбросы которых подлежат регулированию планируемыми (или внедренными)  воздухоохранными мероприятиями',merge_format)
worksheet.merge_range(row_9,2,row_9,3,'Уточненный перечень веществ, причиняющих вред здоровью, выбросы которых подлежат   обязательному регулированию ',merge_format)
worksheet.merge_range(row_9,4,row_9,5,'Окончательный список веществ, выбросы которых подлежат обязательному регулированию',merge_format)
for i in range(len(decrease_elements)):
    worksheet.merge_range(row_9+1+i,2,row_9+1+i,3,table_9.iloc[i,2],main_format)
    worksheet.merge_range(row_9+1+i,4,row_9+1+i,5,table_9.iloc[i,4],main_format)
for i in range(len(names_elements_plan)):
    worksheet.merge_range(row_9+1+i,0,row_9+1+i,1,table_9.iloc[i,0],main_format)
if len(decrease_elements) > len(names_elements_plan):
    for i in range(len(names_elements_plan),len(decrease_elements)):
        worksheet.merge_range(row_9+1+i,0,row_9+1+i,1,table_9.iloc[i,0],main_format)
else:
    for i in range(len(decrease_elements),len(names_elements_plan)):
        worksheet.merge_range(row_9+1+i,2,row_9+1+i,3,table_9.iloc[i,2],main_format)
        worksheet.merge_range(row_9+1+i,4,row_9+1+i,5,table_9.iloc[i,4],main_format)
worksheet.set_row(row_9-1, 60)
#Table 10

worksheet.write(row_10-1,0,'Таблица 13', table_format)
worksheet.merge_range(row_10-1,1,row_10-1,5,'Аэрогенная нагрузка вещества, приходящаяся на население в зоне под  экспозицией', names_format)
worksheet.merge_range(row_10,0,row_10,2,'Наименование вещества',merge_format)
worksheet.merge_range(row_10,3,row_10,5,'Аэрогенная нагрузка вещества (Vi), т/год',merge_format)
for i in range(n_of_elements+1):
    worksheet.merge_range(row_10+1+i,0,row_10+1+i,2,table_10.iloc[i,0],main_format)
    worksheet.merge_range(row_10+1+i,3,row_10+i+1,5,table_10.iloc[i,3],main_format)
worksheet.set_row(row_10-1, 60)
#Table 11

worksheet.write(row_11-3,0,'Таблица 14',table_format)
worksheet.merge_range(row_11-3,1,row_11-3,n_of_factory*3+2,'Планируемая к регулированию и верифицированная аэрогенная нагрузка вещества, создаваемая хозяйствующим субъектом в зоне его размещения',names_format)    
worksheet.merge_range(row_11-2,0,row_11,2,'Вещество',merge_format)
worksheet.merge_range(row_11-2,3,row_11-2,n_of_factory*3+2,'Аэрогенная нагрузка вещества, создаваемая хозяйствующим субъектом в зоне его размещения, т/год',merge_format)

for i in range(n_of_factory):
    worksheet.merge_range(row_11-1,3+3*i,row_11-1,5+3*i, factory_name[i],merge_format)
    worksheet.write(row_11,3*i+3,'планируемая')
    worksheet.write(row_11,3*i+4,'верифицированная')
    worksheet.write(row_11,3*i+5,'расхождение, %')

for i in range(n_of_elements):
    worksheet.merge_range(row_11+1+i,0,row_11+1+i,2,table_11.iloc[i,0],main_format)
worksheet.set_row(row_11-3, 60)
#Table 12

worksheet.write(row_12-2,0,'Таблица 15', table_format)
worksheet.merge_range(row_12-2,1,row_12-2,n_of_elements,'Аэрогенная нагрузка вещества от каждого хозяйствующего субъекта, потенциально достаточная для митигации риска и причиненного вреда здоровью, в зоне под экспозицией',names_format)
worksheet.merge_range(row_12-1,0,row_12,0,'Хозяйствующий субъект',merge_format)
worksheet.merge_range(row_12-1,1,row_12-1,n_of_elements,'Аэрогенная нагрузка вещества, потенциально достаточная для митигации  риска  и причиненного вреда здоровью, в зоне под экспозицией, т/год',merge_format)
for i in range(n_of_elements):
    worksheet.write(row_12,1+i,names_elements[i],merge_format)
worksheet.set_row(row_12-2, 60)
#Table 13

worksheet.write(row_13-2,0,'Таблица 16', table_format)
worksheet.merge_range(row_13-2,1,row_13-2,n_of_elements,'Хроническая средняя суточная доза вещества, не создающая потенциальной опасности причинения вреда здоровью экспонированных лиц (для наихудших сочетанных воздействий в НМУ)',names_format)
worksheet.merge_range(row_13-1,0,row_13,0,'Хозяйствующий субъект',merge_format)
worksheet.merge_range(row_13-1,1,row_13-1,n_of_elements,'Хроническая средняя суточная доза вещества после проведения запланированных субьектом  воздухоохранных мероприятий, не создающая потенциальной опасности причинения вреда здоровью, мг/(кг*сутки)',merge_format)
for i in range(n_of_elements):
    worksheet.write(row_13,1+i,names_elements[i],merge_format)
worksheet.set_row(row_13-2, 60)
#Table 14

worksheet.write(row_14-2,0,'Таблица 17', table_format)
worksheet.merge_range(row_14-2,1,row_14-2,6,'Ранжирование хронической средней суточной дозы вещества, потенциально опасной по причинению вреда здоровью по всему спектру выявленных классов болезней у экспонированных лиц, до и после проведения воздухоохранных мероприятий',names_format)
worksheet.merge_range(row_14-1,0,row_14,0,'Наименование вещества',merge_format)
worksheet.merge_range(row_14-1,1,row_14-1,6,'Ранг хронической средней суточной дозы вещества по степени потенциальной опасности причинения вреда здоровью (Rg)',merge_format)
worksheet.merge_range(row_14,1,row_14,3,'до проведения воздухоохранных мероприятий',merge_format)
worksheet.merge_range(row_14,4,row_14,6,'после проведения воздухоохранных мероприятий',merge_format)
for i in range(n_of_elements):
    worksheet.merge_range(row_14+1+i,1,row_14+i+1,3,table_14.iloc[i,0],main_format)
    worksheet.merge_range(row_14+1+i,4,row_14+i+1,6,table_14.iloc[i,3],main_format)
    worksheet.write(row_14+i+1,0,names_elements[i],main_format)
worksheet.set_row(row_14-2, 60)
#Table 15
    
worksheet.write(row_15-2,0,'Таблица 18', table_format)
worksheet.merge_range(row_15-2,1,row_15-2,n_of_elements,'Масса выброса вещества в атмосферный воздух, на величину которой требуется дополнительное снижение по критерию уменьшения вреда здоровью населения',names_format)
worksheet.merge_range(row_15-1,0,row_15,0,'Хозяйствующий субъект',merge_format)
worksheet.merge_range(row_15-1,1,row_15-1,n_of_elements,'Масса выброса вещества в атмосферный воздух, на величину которой требуется дополнительное снижение по критерию уменьшения вреда здоровью населения , т/год',merge_format)
for i in range(n_of_elements):
    worksheet.write(row_15,1+i,names_elements[i],merge_format)
worksheet.set_row(row_15-2, 60)
#Table 16

worksheet.write(row_16-1,0,'Таблица 20', table_format)
worksheet.merge_range(row_16-1,1,row_16-1,5,'Совокупный причиненный вред здоровью в виде негативного эффекта (по конкретному классу болезни),детерминированного аэрогенной экспозицией после проведения воздухоохранных мероприятий',names_format)
worksheet.merge_range(row_16,0,row_16,2,'Наименование класса болезней',merge_format)
worksheet.merge_range(row_16,3,row_16,5,'Величина совокупного причиненного вреда здоровью, детерминированного аэрогенной экспозицией после проведения воздухоохранных мероприятий',merge_format)
for i in range(n_of_disease):
    worksheet.merge_range(row_16+1+i,0,row_16+1+i,2,names_disease[i],main_format)
    worksheet.merge_range(row_16+1+i,3,row_16+1+i,5,table_16.iloc[i,3],main_format)
worksheet.set_row(row_16-1, 60)
#Table 17
    
worksheet.write(row_17-2,0,'Таблица 22', table_format)
worksheet.merge_range(row_17-2,1,row_17-2,7,'Эффективность запланированных (внедренных) воздухоохранных мероприятий по критериям предотвращенного вреда здоровью ',names_format)
worksheet.merge_range(row_17-1,0,row_17-1,3,'Совокупный причиненный вред, детерминированный аэрогенной экспозицией ',merge_format)
worksheet.merge_range(row_17,0,row_17,1,'до проведения воздухоохранных мероприятий',merge_format)
worksheet.merge_range(row_17,2,row_17,3,'после проведения воздухоохранных мероприятий',merge_format)
worksheet.merge_range(row_17-1,4,row_17,5,'Эффективность (Э), %',merge_format)
worksheet.merge_range(row_17-1,6,row_17,7,'Степень эффективности',merge_format)
for i in range(4):
    worksheet.merge_range(row_17+1,2*i,row_17+1,2*i+1,table_17.iloc[0,2*i],main_format)
worksheet.set_row(row_17-2, 60)
#Table 18

worksheet.write(row_18-1,0,'Таблица 23', table_format)
worksheet.merge_range(row_18-1,1,row_18-1,4,'Дополнительно к запланированному объём снижения выброса вещества для достижения эффективности воздухоохранных мероприятий, %',names_format)
worksheet.merge_range(row_18,0,row_18,1,'Вещество',merge_format)
worksheet.merge_range(row_18,2,row_18,4,'Дополнительно к запланированному объём снижения выброса',merge_format)
for i in range(n_of_elements):
    worksheet.merge_range(row_18+1+i,0,row_18+1+i,1,names_elements[i],merge_format)
    #worksheet.merge_range(row_18+1+i,2,row_18+1+i,4,addit_decrease[i],merge_format)
#for i in range(n_of_elements):
#    if addit_decrease[i] >= 50:
#        worksheet.merge_range(row_18+1+i,0,row_18+1+i,1,names_elements[i],danger_format_5)

#General
worksheet.set_column(0, 100, 15)

pd.set_option('display.max_colwidth', 0)
print("--- %s seconds ---" % (time.time() - start_time))
writer.save()