# -*- coding: utf-8 -*-
"""
Created on Sat Oct  6 11:06:59 2018

@author: roberto.valdez
"""

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import numpy as np
import os
import sys
import operator as op
from openpyxl import load_workbook


class Persona:

    # Class Attributes
    errores = ''    

    # Initializer / Instance Attributes
    def __init__(self, nombre, cedula, sexo, edad, ts1, ts2, cu, pos):
        self.nombre = str(nombre)
        self.cedula = str(cedula)
        self.sexo = sexo
        self.edad = edad
        self.ts1 = ts1
        self.ts2 = ts2
        self.cu = str(cu)  #Codigo único
        self.pos = pos
        
    def b_search(self, x):
        pos = -1
        b = 0
        e = len(x) - 1
        m = int((b+e)/2)    
        while b <= e and pos == -1:            
            if x[m].nombre == self.nombre:
                pos = m 
            else:
                if x[m].nombre > self.nombre:
                    e = m - 1
                else:
                    b = m + 1
            m = int((b+e)/2) 
        return(pos)
        
        
    def comparador(self, lst_persona):
        pos = self.b_search(lst_persona)
        #Aca se puede trabajar en otros métodos de búsqueda
        return(pos)



class Application(ttk.Frame):
    
    def __init__(self, main_window):        
        super().__init__(main_window)         
        main_window.title("Procesador Impuestos Diferidos (beta1)")
        
        self.dataload_button = ttk.Button(
            self, text="Cargar Archivo", command=self.load_file)
        self.dataload_button.place(x=30, y=100)
        
        self.run_button = ttk.Button(
            self, text="Ejecutar", command=self.run_diag)
        self.run_button.place(x=30, y=140)
        
        self.progressbar = ttk.Progressbar(self, orient="horizontal",
                                        length=200, mode="determinate")
        self.progressbar.place(x=30, y=180, width=200)
        
                
        self.text1 = tk.StringVar()
        self.text1.set('[.]')
        self.text2 = tk.StringVar()
        self.text2.set("%s%%" % '{:.2f}'.format(0.00))
        self.text3 = tk.StringVar()
        self.text3.set('...')        
        
        self.label1 = tk.Label(self, textvariable = self.text1)
        self.label1.place(x=120,y=100)
        self.label2 = tk.Label(self, textvariable = self.text2)
        self.label2.place(x=235,y=180) 
        self.label3 = tk.Label(self, textvariable = self.text3)
        self.label3.place(x=30,y=205) 
        
        
        
        if getattr(sys, 'frozen', False):
            # frozen
            mdir = os.path.dirname(sys.executable)
        else:
            # unfrozen
            mdir = os.path.dirname(os.path.realpath(__file__))
        
        os.chdir(mdir)
        
        self.logo = tk.PhotoImage(file='logo.gif') 
        self.label4 = tk.Label(self, image = self.logo)
        self.label4.place(x=190,y=5)        

        self.workbook = {}  
        self.sheet_names = []
        self.data_prob = False
        self.data_loaded = False
        self.totalIter = 0
        self.currentIter = 0
        
        
        self.place(width=340, height=250)        
        main_window.geometry("350x250")
        
    
    def start_pb(self, d_length):
        self.progressbar["value"] = 0
        self.progressbar["maximum"] = d_length 
        
    
        
    def load_file(self):             
        self.fileName = filedialog.askopenfilename(filetypes=[("xlsx files","*.xlsx")])       
        if self.fileName != "":
            self.text1.set('Cargando Archivo.....')  
            main_window.update_idletasks() 
            try:                
                self.workbook = pd.read_excel(self.fileName, sheetname=None)
                self.sheet_names = list(self.workbook.keys())   
                self.sheet_names.sort(key = int, reverse = True)   
                
                for names in self.sheet_names:
                    df = self.workbook[names]
                    if len(df.columns) != 19:   #set to be 19 columns in data
                        self.workbook = {}  
                        self.text1.set('[.]') 
                        answer = messagebox.askokcancel("Procesador Impuestos Diferidos", 'Error en columnas en: ' + names)
                        if answer or not answer:
                            main_window.destroy()                        
                            self.data_prob = True
                            break
                        
                    try:
                        df['ts2'] = np.array(df['ts2']).astype(float)
                    except ValueError:
                        answer = messagebox.askokcancel("Procesador Impuestos Diferidos", 'Error en columna ts2 en: ' + names)
                        if answer or not answer:
                            main_window.destroy()                                            
                            self.data_prob = True
                            break
                    
                       
                    #Se debe hacer mas validaciones                        
                if not self.data_prob:                
                    self.data_loaded = True
                    self.totalIter = len(self.sheet_names)*4 + 4
                    self.start_pb(self.totalIter)  #aca hay que aumentar para el número de pasos que hago
                    self.text1.set('[Archivo Cargado]')
                    
            except Exception as e:                
                answer = messagebox.askokcancel("Procesador Imp. Diferidos", e)
                if answer or not answer:
                    main_window.destroy()
                
    
    def get_classPers(self, sheet_names, workbook):
        #leer la base y transformar a todos los ingresos en la data en objetos de la clase personas
        personas = []
    
        for names in sheet_names:
            df = workbook[names]
            pers_anio = []    
            i = 1  
            for index, row in df.iterrows():
                if names == sheet_names[0]:             
                    pers_anio.append(Persona(row[2], row[3], row[4], row[5], row[6], row[7], i, i-1))            
                else:
                    pers_anio.append(Persona(row[2], row[3], row[4], row[5], row[6], row[7], 'comp', i-1))        
                i+=1        
            personas.append(pers_anio)
            
            self.progressbar["value"] += 1
            self.currentIter += 1
            perc = '{:.2f}'.format(round(self.currentIter/self.totalIter,4)*100)
            print("%s%%" % perc)
            self.text2.set("%s%%" % perc) 
            main_window.update_idletasks() 
            
            
        return(personas)
        
    def get_compPers(self, sheet_names, personas):
        #Retornar la clase personas con las personas de los otros años asignadas su código único
        pers_comp = []
    
        pers_comp.append(personas[0])
        
        for i in range(len(sheet_names)-1):
            pivot = pers_comp[i]
            toComp = personas[i+1]
            
            std_toComp = sorted(toComp, key=op.attrgetter('nombre'))
            gt1 = list(filter(lambda x: x.ts2 >= 1 and x.cu != 'comp', pivot))
            lt1 = list(filter(lambda x: x.ts2 < 1 or x.cu == 'comp', pivot))
            
            for obj in gt1:
                pos = obj.comparador(std_toComp)
                if pos >= 0:
                    std_toComp[pos].cu = obj.cu
                else:           
                    obj.errores = 'No aparece en ' + sheet_names[i+1]
            
            pivot00 = sorted(gt1 + lt1, key=op.attrgetter('pos'))
            toComp11 = sorted(std_toComp, key=op.attrgetter('pos'))  
            
            pers_comp[i] = pivot00
            pers_comp.append(toComp11)
            
            self.progressbar["value"] += 1
            self.currentIter += 1
            perc = '{:.2f}'.format(round(self.currentIter/self.totalIter,4)*100)
            print("%s%%" % perc)
            self.text2.set("%s%%" % perc) 
            main_window.update_idletasks()
        
        return(pers_comp)  
        
    def get_pdBase(self, sheet_names, workbook, pers_comp):
        #add colum in each database named CU, then we use pandas to merge by that cu
        #we return a database merged
    
        for i in range(len(sheet_names)):
            anioi = pers_comp[i]
            cu = np.array([o.cu for o in anioi])  #Aca se podria trabajar en duplicados antes de pasarle a la base de datos
            attr = np.array([o.errores for o in anioi])
            cols = list(workbook[sheet_names[i]])
            cols = [s + '_' + sheet_names[i] for s in cols]
            workbook[sheet_names[i]].columns = cols
            workbook[sheet_names[i]]['CU'] = cu
            workbook[sheet_names[i]]['Errores'+ '_' + sheet_names[i]] = attr
            
            self.progressbar["value"] += 1
            self.currentIter += 1
            perc = '{:.2f}'.format(round(self.currentIter/self.totalIter,4)*100)
            print("%s%%" % perc)
            self.text2.set("%s%%" % perc) 
            main_window.update_idletasks()
        
        
        result = workbook[sheet_names[0]]
        for i in range(len(sheet_names)-1):
            data1 = workbook[sheet_names[i+1]]
            result = pd.merge(left=result, right=data1, on="CU", how="left")
            
            self.progressbar["value"] += 1
            self.currentIter += 1
            perc = '{:.2f}'.format(round(self.currentIter/self.totalIter,4)*100)
            print("%s%%" % perc)
            self.text2.set("%s%%" % perc) 
            main_window.update_idletasks()
        
        return(result)     
    
    def write_File(self, result):    
        #prepare database for presentation and write it to excel (or csv) file
    
        #result.to_csv("Result.csv")           
        book = load_workbook('Existing_File.xlsx')
        
        self.progressbar["value"] += 1
        self.currentIter += 1
        perc = '{:.2f}'.format(round(self.currentIter/self.totalIter,4)*100)
        print("%s%%" % perc)
        self.text2.set("%s%%" % perc) 
        main_window.update_idletasks()        
        
        writer = pd.ExcelWriter(self.fileName[0:-5]+"_RESULT.xlsx", engine='openpyxl') 
        
        self.progressbar["value"] += 1
        self.currentIter += 1
        perc = '{:.2f}'.format(round(self.currentIter/self.totalIter,4)*100)
        print("%s%%" % perc)
        self.text2.set("%s%%" % perc) 
        main_window.update_idletasks()
        
        
        writer.book = book
        
        self.progressbar["value"] += 1
        self.currentIter += 1
        perc = '{:.2f}'.format(round(self.currentIter/self.totalIter,4)*100)
        print("%s%%" % perc)
        self.text2.set("%s%%" % perc) 
        main_window.update_idletasks()        
        
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        
        self.progressbar["value"] += 1
        self.currentIter += 1
        perc = '{:.2f}'.format(round(self.currentIter/self.totalIter,4)*100)
        print("%s%%" % perc)
        self.text2.set("%s%%" % perc) 
        main_window.update_idletasks()
    
        result.to_excel(writer, "Hoja1", header=True, index=False)
        
        self.progressbar["value"] += 1
        self.currentIter += 1
        perc = '{:.2f}'.format(round(self.currentIter/self.totalIter,4)*100)
        print("%s%%" % perc)
        self.text2.set("%s%%" % perc) 
        main_window.update_idletasks()
    
        writer.save()   
    
        self.progressbar["value"] += 1
        self.currentIter += 1
        perc = '{:.2f}'.format(round(self.currentIter/self.totalIter,4)*100)
        print("%s%%" % perc)
        self.text2.set("%s%%" % perc) 
        main_window.update_idletasks()
        
        
    def run_diag(self):
        if self.data_loaded:
            try:
                               
                self.text3.set('Paso 1 de 4: Cargando bases (objetos)')
                main_window.update_idletasks() 
                personas = self.get_classPers(self.sheet_names, self.workbook)
                print('Paso 1 de 4: Cargando bases (objetos)')
                
                self.text3.set('Paso 2 de 4: Comparando')
                main_window.update_idletasks() 
                pers_comp = self.get_compPers(self.sheet_names, personas) 
                print('Paso 2 de 4: Comparando')
                
                self.text3.set('Paso 3 de 4: Arreglando Base result.')
                main_window.update_idletasks() 
                result = self.get_pdBase(self.sheet_names, self.workbook, pers_comp)
                print('Paso 3 de 4: Arreglando Base result.')
                     
                self.text3.set('Paso 4 de 4: Guardando archivo final')
                main_window.update_idletasks() 
                self.write_File(result)  
                print('Paso 4 de 4: Guardando archivo final')
                
                answer = messagebox.askokcancel("Procesador Imp. Diferidos", "El proceso concluyó satisfactoriamente")
                if answer:
                    main_window.destroy()
                
            except Exception as e:                
                answer = messagebox.askokcancel("Procesador Imp. Diferidos", e)
                if answer:
                    main_window.destroy()          
            
        else:
           messagebox.askokcancel("Procesador Impuestos Diferidos", 'Por favor cargar un archivo válido')
                       
        
main_window = tk.Tk()
app = Application(main_window)
app.mainloop()       
        
        
        