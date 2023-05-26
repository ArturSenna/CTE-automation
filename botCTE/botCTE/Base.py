from emissions import *
from ttkthemes import ThemedStyle
from tkcalendar import DateEntry
import ctypes.wintypes
from tkinter import filedialog as fd
from tkinter import Toplevel
from tkinter import ttk
from tkinter import *
import threading
from json import loads
import csv
import pandas as pd
import numpy as np
import requests
import time
import bot
from functions import *

# Create Object
root = Tk()
root.title('CTe LogLife')
root.geometry("")
root.resizable(False, False)
root.iconbitmap('my_icon.ico')
thread_0 = Start(root)
thread_1 = Start(root)

# Setting tabs

tabs = ttk.Notebook(root)
tab1 = ttk.Frame(tabs)
tab2 = ttk.Frame(tabs)
tab3 = ttk.Frame(tabs)
tab4 = ttk.Frame(tabs)

tabs.add(tab1, text='CTe')
tabs.add(tab2, text='CTe complementar')
tabs.add(tab4, text='Pastas Relatório')
tabs.add(tab3, text='Pasta CTE')

tabs.pack(expand=1, fill="both")

tab2_frame = Frame(tab2, pady=22)
tab2_frame.pack()

tab2_frame2 = Frame(tab2)
tab2_frame2.pack()

tab4_frame = Frame(tab4)
tab4_frame.pack()

# Set calendar

today = dt.datetime.today()

dia = today.day
mes = today.month
ano = today.year

style = ThemedStyle(root)
style.theme_use('breeze')

cal_config = {'selectmode': 'day',
              'day': dia,
              'month': mes,
              'year': ano,
              'locale': 'pt_BR',
              'firstweekday': 'sunday',
              'showweeknumbers': False,
              'bordercolor': "white",
              'background': "white",
              'disabledbackground': "white",
              'headersbackground': "white",
              'normalbackground': "white",
              'normalforeground': 'black',
              'headersforeground': 'black',
              'selectbackground': '#00a5e7',
              'selectforeground': 'white',
              'weekendbackground': 'white',
              'weekendforeground': 'black',
              'othermonthforeground': 'black',
              'othermonthbackground': '#E8E8E8',
              'othermonthweforeground': 'black',
              'othermonthwebackground': '#E8E8E8',
              'foreground': "black"}

cal = DateEntry(tab1, **cal_config)
cal1 = DateEntry(tab2_frame, **cal_config)
cal2 = DateEntry(tab2_frame, **cal_config)

cal.grid(column=1, row=0, padx=30, pady=10, sticky="E, W")
cal1.grid(column=1, row=0, padx=30, pady=10, sticky="E, W")
cal2.grid(column=1, row=1, padx=30, pady=10, sticky="E, W")

# Read file name

cte_folder = StringVar()

try:
    with open('filename.txt') as m:
        text = m.read()
    lines = text.split('\n')
    cte_folder.set(lines[0])
except FileNotFoundError:
    cte_folder.set('Pasta onde os arquivos CTe estão.')

cte_Folder_Label = ttk.Label(tab3, width=20, text=cte_folder.get(), wraplength=140)
cte_Folder_Label.grid(column=1, row=0, padx=5, pady=5, ipadx=8)

# Read folder path

folderpath = StringVar()

try:
    with open('folderpath.txt') as m:
        text = m.read()
    lines = text.split('\n')
    folderpath.set(lines[0])
except FileNotFoundError:
    folderpath.set('Pasta para CTe normal')

folder_Label = ttk.Label(tab4_frame, width=20, text=folderpath.get(), wraplength=140)
folder_Label.grid(column=1, row=0, sticky="E, W", padx=20, pady=10)

folderpath2 = StringVar()

try:
    with open('folderpath2.txt') as m:
        text = m.read()
    lines = text.split('\n')
    folderpath2.set(lines[0])
except FileNotFoundError:
    folderpath2.set('Pasta para CTe complementar')

folder_Label2 = ttk.Label(tab4_frame, width=20, text=folderpath2.get(), wraplength=140)
folder_Label2.grid(column=1, row=1, sticky="E, W", padx=20, pady=10)

browse1 = Browse(cte_Folder_Label)
browse2 = Browse(folder_Label)
browse3 = Browse(folder_Label2)

# Add Buttons
ttk.Button(tab1,
           text="Emitir lista de CTe",
           command=lambda: thread_0.start_thread(
               cte_list, progressbar, arguments=[cal.get_date(),
                                                 folderpath.get(),
                                                 cte_folder.get(),
                                                 root]
           )).grid(column=2, row=0, padx=10, pady=15, sticky="N, S, E, W")

ttk.Button(tab1,
           text="Emitir",
           command=lambda: thread_0.start_thread(
               cte_unique, progressbar, arguments=[cal.get_date(),
                                                   folderpath.get(),
                                                   cte_folder.get(),
                                                   cte_type.get(),
                                                   cte_s.get(),
                                                   volumes.get(),
                                                   root]
           )).grid(column=2, row=1, padx=10, pady=10, sticky="N, S, E, W")

ttk.Button(tab2_frame,
           text="Emitir",
           command=lambda: thread_1.start_thread(
               cte_complimentary, progressbar2, arguments=[cal1.get_date(),
                                                           cal2.get_date(),
                                                           folderpath2.get(),
                                                           cte_folder.get(),
                                                           root]
           )).grid(column=2, row=0, rowspan=2, padx=10, pady=15, ipady=15, sticky=" E, W")

ttk.Button(tab2_frame2,
           text="Emitir",
           command=lambda: thread_1.start_thread(
               cte_complimentary, progressbar2, arguments=[cal1.get_date(),
                                                           cal2.get_date(),
                                                           folderpath2.get(),
                                                           cte_folder.get(),
                                                           root,
                                                           True,
                                                           cte_cs.get()]
           )).grid(column=2, row=0, padx=10, pady=10, sticky="N, S, E, W")

ttk.Button(tab3,
           text="Procurar",
           command=lambda: browse1.browse_folder(cte_folder, 'filename.txt', master=tab3,
                                                 label_config={'wraplength': 140, 'width': 20},
                                                 grid_config={'column': 1, 'row': 0, 'sticky': "E, W", 'padx': 20,
                                                              'pady': 10})
           ).grid(column=2, row=0, padx=5, pady=5, sticky="E, W")

ttk.Button(tab4_frame,
           text="Procurar",
           command=lambda: browse2.browse_folder(folderpath, 'folderpath.txt', master=tab4_frame,
                                                 label_config={'wraplength': 140, 'width': 20},
                                                 grid_config={'column': 1, 'row': 0, 'sticky': "E, W", 'padx': 20,
                                                              'pady': 10})
           ).grid(column=2, row=0, padx=5, pady=5, sticky="E, W")

ttk.Button(tab4_frame,
           text="Procurar",
           command=lambda: browse3.browse_folder(folderpath2, 'folderpath2.txt', master=tab4_frame,
                                                 label_config={'wraplength': 140, 'width': 20},
                                                 grid_config={'column': 1, 'row': 1, 'sticky': "E, W", 'padx': 20,
                                                              'pady': 10})
           ).grid(column=2, row=1, padx=5, pady=5, sticky="E, W")

# Labels

ttk.Label(tab1, text="Data da Emissão:").grid(column=0, row=0, padx=10, pady=10, sticky='W, E')
ttk.Label(tab1, text="CTe avulso:").grid(column=0, row=1, padx=10, pady=10, sticky='W, E')
ttk.Label(tab1, text="Vols:").grid(column=1, row=4, padx=10, pady=10, sticky='W, E')

ttk.Label(tab2_frame, text="Data Inicial:").grid(column=0, row=0, padx=10, pady=10, sticky='W, E')
ttk.Label(tab2_frame, text="Data Final:").grid(column=0, row=1, padx=10, pady=10, sticky='W, E')

ttk.Label(tab2_frame2, text="CTe avulso:").grid(column=0, row=0, padx=10, pady=10, sticky='W, E')

ttk.Label(tab3, text="Pasta CTE:").grid(column=0, row=0, padx=5, pady=5, ipadx=10)

ttk.Label(tab4_frame, text="CTe normal:").grid(column=0, row=0, padx=5, pady=5)
ttk.Label(tab4_frame, text="CTe complementar:").grid(column=0, row=1, padx=5, pady=5)

# Adding Entry texts

cte_s = StringVar()

ttk.Entry(tab1,
          textvariable=cte_s, width=14).grid(column=1, row=1, padx=10, pady=10)

cte_cs = StringVar()

ttk.Entry(tab2_frame2,
          textvariable=cte_cs, width=14).grid(column=1, row=0, padx=10, pady=10)

volumes = StringVar()
ttk.Entry(tab1,
          textvariable=volumes, width=8).grid(column=1, row=4, pady=5, ipadx=1)
volumes.set('1')

# Radio Buttons

cte_type = IntVar()

ttk.Radiobutton(tab1, text="Normal", value=0, variable=cte_type).grid(column=0, row=4, ipadx=18)
ttk.Radiobutton(tab1, text="Simbólico", value=1, variable=cte_type).grid(column=0, row=5, ipadx=10)

# Progress Bar

progressbar = ttk.Progressbar(tab1, mode='indeterminate')
progressbar.grid(column=0, row=6, sticky='W, E', columnspan=3)

progressbar2 = ttk.Progressbar(tab2, mode='indeterminate')
progressbar2.pack(side=BOTTOM, fill='x')

# Auto resize tabs

tab1.rowconfigure(2, weight=2)
tab1.columnconfigure(0, weight=1)
tab1.columnconfigure(1, weight=1)
tab1.columnconfigure(2, weight=1)

tab3.rowconfigure(0, weight=1)
tab3.columnconfigure(0, weight=1)
tab3.columnconfigure(1, weight=1)
tab3.columnconfigure(2, weight=1)

# Execute Tkinter
root.mainloop()
