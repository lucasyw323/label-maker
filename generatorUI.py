from datetime import date
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu

import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog

import pandas as pd
import numpy as np
import os

from generateLabels import *
from generateCoC import *
from generatePackingList import *
from generateDocs import *

class Page(tk.Frame):
    def __init__(self, container, app):
        tk.Frame.__init__(self, container)
        self.winfo_toplevel().title("Labelmaker")

        self.app = app

        coc_template = tk.Label(self, text='CoC Template:')
        coc_template.grid(row=0, column=0, padx = 10, pady = 10)
        choose_coc = ttk.Button(self, text = "Browse...", command = 
                           lambda: self.getFile('CoC Template', '*.docx'))
        choose_coc.grid(row=0, column=1, pady = 10)
        self.coc_file = ttk.Label(self, text = "No file chosen")
        self.coc_file.grid(row=0, column=2, pady = 10)

        plist_template = tk.Label(self, text='Packing List Template:')
        plist_template.grid(row=1, column=0, padx = 10)
        choose_plist = ttk.Button(self, text = "Browse...", command = 
                           lambda: self.getFile('Packing List Template', 
                                                '*.docx'))
        choose_plist.grid(row=1, column=1, pady = 10)
        self.plist_file = ttk.Label(self, text = "No file chosen")
        self.plist_file.grid(row=1, column=2, pady=10)

        sheet = tk.Label(self, text='Input sheet:')
        sheet.grid(row=2, column=0, padx = 10, pady = 10)
        choose_sheet = ttk.Button(self, text = "Browse...", command = 
                           lambda: self.getFile('Input sheet', '*.xlsx'))
        choose_sheet.grid(row=2, column=1, pady = 10)
        self.sheet_file = ttk.Label(self, text = "No file chosen")
        self.sheet_file.grid(row=2, column=2, pady = 10)

        dest = tk.Label(self, text='Destination folder:')
        dest.grid(row=3, column=0, padx = 10, pady = 10)
        choose_dest = ttk.Button(self, text = "Browse...", command = 
                           self.getFolder)
        choose_dest.grid(row=3, column=1, pady = 10)
        self.dest_dir = ttk.Label(self, text = "No folder chosen")
        self.dest_dir.grid(row=3, column=2, pady = 10)

        generate = ttk.Button(self, text = "Generate\nDocuments", command = 
                           self.app.generate)
        generate.grid(row = 4, column = 0, padx = 10, pady = 10, columnspan=3)

        self.label = ttk.Label(self, text = "")
        self.label.grid(row = 5, column = 0, padx = 10, pady = 10, columnspan=3)

        # preset = ttk.Button(self, text = "preset", command = 
        #                    self.app.preset)
        # preset.grid(row = 6, column = 0, padx = 10, pady = 10, columnspan=3)
    
    def getFile(self, doc, filetype):
        file_prompt = ''
        if(filetype == '*.xlsx'):
            file_prompt = 'Excel Workbook'
        elif(filetype == '*.docx'):
            file_prompt = 'Word Document'
        file = filedialog.askopenfilename(initialdir = "/",title = "Select " +
                                            doc, filetypes = ((file_prompt,
                                            filetype),("all files", "*.*")))
        if(doc == "Input sheet"):
            self.sheet_file['text'] = file
            app.book = file
        elif(doc == "CoC Template"):
            self.coc_file['text'] = file
            app.coc_template = file
        elif(doc == "Packing List Template"):
            self.plist_file['text'] = file
            app.plist_template = file
        return 
    
    def getFolder(self):
        folder = filedialog.askdirectory(initialdir = "/",title = "Select folder")
        self.dest_dir['text'] = folder
        self.app.dest = folder
        return 

class Application(tk.Tk):
    def __init__(self, *args, **kwargs):
        # __init__ function for class Tk
        tk.Tk.__init__(self, *args, **kwargs)
            
        # creating a container
        container = tk.Frame(self) 
        container.pack(side = "top", fill = "both", expand = True)
    
        container.grid_rowconfigure(0, weight = 1)
        container.grid_columnconfigure(0, weight = 1)
    
        # initializing frames to an empty array
        self.frames = {}
        
        page = Page(container, self)
        self.frames[Page] = page  
        page.grid(row = 0, column = 0, sticky ="nsew")
        
        self.book = ""
        self.coc_template = ""
        self.plist_template = ""
        self.dest = ""
        
        self.frames[Page].tkraise()
        self.update()

    def preset(self):
        self.book = 'C:/Lucas/TreadStone/label-maker/Empty Input Template.xlsx'
        self.plist_template = 'C:/Lucas/TreadStone/label-maker/Templates/Packing List Template New.docx'
        self.coc_template = 'C:/Lucas/TreadStone/label-maker/Templates/CoC Template New.docx'
        self.dest = 'C:/Lucas/TreadStone/testing'
        return

    def generate(self):
        if(self.book == '' or self.coc_template == '' or 
           self.plist_template == '' or self.dest == ''):
            self.frames[Page].label['text'] = 'Missing file/folder'
            return
        # print(self.book, self.plist_template, self.coc_template, 
        #              self.dest)
        loc = generateDocs(self.book, self.plist_template, self.coc_template, 
                     self.dest, self.frames[Page].label)
        if(loc == 404):
            return
        self.frames[Page].label['text'] = ('Documents saved to\n' + loc)

app = Application()

app.mainloop()