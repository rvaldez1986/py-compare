# -*- coding: utf-8 -*-
"""
Created on Sat Oct  6 11:06:59 2018

@author: roberto.valdez
"""

import tkinter as tk
from tkinter import (Tk, ttk, filedialog, messagebox, NORMAL, DISABLED)
import multiprocessing as mp
from openpyxl import load_workbook
import operator as op
import numpy as np
import pandas as pd
import re              
import unicodedata
import os
import sys
from difflib import SequenceMatcher
import csv


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()
    
def pwCount(a, b):
    als = a.split()
    bls = b.split()
    pct = len([x for x in als if x in bls])/max(len(als),1)
    return(pct)


class Persona:

    # Class Attributes
    errores = ''    

    # Initializer / Instance Attributes
    def __init__(self, nombre, cedula, sexo, edad, ts1, ts2, cod_emp, cu, pos):
        self.nombre = str(nombre)
        self.cedula = str(cedula)
        self.sexo = sexo
        self.edad = edad
        self.ts1 = ts1
        self.ts2 = ts2
        self.cu = str(cu)  #Codigo único
        self.pos = pos
        self.cod_emp = int(cod_emp)
        
    def b_search(self, x, attr):
        pos = -1
        b = 0
        e = len(x) - 1
        m = int((b+e)/2)    
        while b <= e and pos == -1:            
            if getattr(x[m], attr) == getattr(self, attr):
                pos = m 
            else:
                if getattr(x[m], attr) > getattr(self, attr):
                    e = m - 1
                else:
                    b = m + 1
            m = int((b+e)/2) 
        return(pos)
        
    def ced_search(self, y):
        pos = -1
        filt1 = list(filter(lambda x: x.cedula != '' , y))
        if len(filt1)>0:
            lst_cedula = sorted(filt1, key=op.attrgetter('cedula'))
            post = self.b_search(lst_cedula, 'cedula')
            if post >= 0:
                namet = lst_cedula[post].nombre
                pos = [ x.nombre for x in y ].index(namet)                
        return(pos)
            
        
    
    def dst_search(self, y):
        pos = -1
        if self.cedula != '':
            filter0 = list(filter(lambda x: x.cedula == '' , y))
            filter1 = list(filter(lambda x: x.edad == self.edad - 1 , filter0))
        else:            
            filter1 = list(filter(lambda x: x.edad == self.edad - 1 , y))        
        
        filterFinal = list(filter(lambda x: x.sexo == self.sexo  , filter1)) 
        
        list1 = [similar(self.nombre, o.nombre) for o in filterFinal]
        if len(list1)>0 and max(list1) >= 0.78:
            pos1 = list1.index(max(list1))
            name1 = filterFinal[pos1].nombre
            pnames = pwCount(self.nombre, name1)
            if pnames >= 0.75:
                pos = [ x.nombre for x in y ].index(name1)     
            
        return(pos)
        
        
        
    def comparador(self, lst_persona):
        #Buscar por nombre
        pos = self.b_search(lst_persona, 'nombre')
       
       #Si no se le encuentra buscar por cedula        
        if pos == -1 and self.cedula != '':
            pos = self.ced_search(lst_persona)            
                    
         #Si no se le encuentra buscar de otra manera            
        if pos == -1:
            pos = self.dst_search(lst_persona)   
        
        return(pos)
        
        
def load_file(fileName):
    workbook = pd.read_excel(fileName, sheetname=None)      
    
    return(workbook)
    
def std_text(text, sheet_name, nro, errores):
    """
    Strip accents from input String.

    :param text: The input string.
    :type text: String.

    :returns: The processed String.
    :rtype: String.
    """
   
    try:         
        if (not text) or (text == None) or pd.isnull(text):
            errores.append('Problemas de nombres vacios en ' + sheet_name + ', posición nro: ' + str(nro))
            return        
        
        text = text.upper().strip(' ,.-$')
        text = re.sub( '\s+', ' ', text )
        
        try:
            text = unicode(text, 'utf-8')
        except (TypeError, NameError): # unicode is a default on python 3 
            pass
        text = unicodedata.normalize('NFD', text)
        text = text.encode('ascii', 'ignore')
        text = text.decode("utf-8")
        text = str(text)
        
        text = re.sub('[^A-Z0-9 ]+', '', text)
        return text
    except :        
        errores.append('Problemas en caracteres de nombres en ' + sheet_name + ', posición nro: ' + str(nro))
        return

def findDup(df, sheet_name, errores):    
    if len(df['nombre']) != len(set(df['nombre'])):        
        errores.append('Problemas duplicados en nombres en ' + sheet_name)
        return    

def ced_val(x):
    x = str(x).strip(' ,.-$')
    if len(x) == 9:
        y = str(0) + x        
    elif len(x) == 10:
        y = x        
    else:
        return('')        
    try:
        z = np.array([int(i) for i in y])     
    except:
        return('')
    
    p = z[:-1][1::2]
    i = z[:-1][0::2]
    ud = z[-1]
    provincia = i[0] * 10 + p[0]    
    
    if provincia < 1 or provincia > 24:
        return('')
    
    i2 = np.where(i>=5,2*i-9,2*i)
    sp = sum(p)
    si = sum(i2)
    stot = sp + si
    DS = int(np.ceil(stot/10)*10)
    vrf = (DS - stot) - ud  
    if vrf == 0:
        return(y)
    else:
        return('')    
        
def val_nbr(num, texto, sheet_name, nro, errores):    
    try:    
        num = float(num)        
        
        if pd.isnull(num):
            errores.append(texto + sheet_name + ', posición nro: ' + str(nro))
            return        
        
        return(num)   
        
    except:        
        errores.append(texto + sheet_name + ', posición nro: ' + str(nro))
        return
    
def val_sx(txt, sheet_name, nro, errores):    
    try:  
        txt = str(txt)
        txt = txt.strip(' ,.-$')
        if txt == 'F' or txt == 'M':     
            return txt
        else:
            errores.append('Problemas en sexo en ' + sheet_name + ', posición nro: ' + str(nro))
            return
    except :        
        errores.append('Problemas en sexo en ' + sheet_name + ', posición nro: ' + str(nro))
        return    
    


def validData(df, sheet_name):

    errores = []    
    
        
    if len(df.columns) != 19:
        errores.append('Problemas en nro de columnas en ' + sheet_name)
        return(df, errores)  
        
    null_columns = list(df.columns[df.isnull().any()])
    
    
    if len(null_columns) > 0:
        myList = ','.join(map(str, null_columns))
        errores.append('Las columnas: ' + myList + ' en: ' + sheet_name + ' tienen valores vacíos')
       
    
    
    df.columns = ['no', 'tipo', 'nombre', 'cedula', 'sexo', 'edad', 'ts1', 'ts2',
                      'tf', 'tw', 'reserva_jub', 'costo_laboral_jub',
                      'interes_neto_jub', 'gasto_jub', 'reserva_des', 'costo_laboral_des', 'interes_neto_des',
                     'gasto_des', 'codigo_empresa']
                     
    df['no'] = list(range(1,(df.shape[0]+1)))   

    df['nombre'] = [std_text(name, sheet_name, nro, errores) for name,nro in zip(df['nombre'], df['no'])]

    findDup(df, sheet_name, errores)

    df['cedula'] = [ced_val(ced) for ced in df['cedula']]
    
        
    df['ts2'] = [val_nbr(num, 'Problemas en ts2 en ', sheet_name, nro, errores) for num,nro in zip(df['ts2'], df['no'])]
    df['edad'] = [val_nbr(num, 'Problemas en edad en ', sheet_name, nro, errores) for num,nro in zip(df['edad'],df['no'])]

    df['sexo'] = [val_sx(txt, sheet_name, nro, errores) for txt,nro in zip(df['sexo'],df['no'])]
    
    df['codigo_empresa'] = [val_nbr(num, 'Problemas en código empresa en ', sheet_name, nro, errores) for num,nro in zip(df['codigo_empresa'],df['no'])]

    errores1 = list(set(errores))
    
    return(df, errores1)
    
def validParam(workbook, sheet_names):
    errores = []
    try:
        df = workbook['Parametros'] 
        df.columns = ['Año', 'ts2_min']
        df['Año'] = [str(value) for value in df['Año']]
        try:
           df['ts2_min'] = [float(value) for value in df['ts2_min']]
           
        except:
           errores.append('Caracteres no numéricos para ts2 min en parámetros')
           ts2Dict = {}
           return(ts2Dict, errores)
        
        ts2Dict = df.set_index('Año').T.to_dict('records')[0]
        
        un_st = list(set([str(int(float(num))) for num in sheet_names]))
        un_st.sort(key=int, reverse = True)
        
        
        dif = [float(numA) - float(numB) for numA,numB in zip(un_st[:-1],un_st[1:])]
       
        
        if max(dif) != 1 or min(dif) != 1:
            errores.append('Existen años faltantes dentro de pestañas (Hay saltos entre años diferentes que 1)')
            ts2Dict = {}
            return(ts2Dict, errores)
            
        
        if not set(ts2Dict.keys()) == set(un_st):            
            errores.append('años en parámetros están incompletos')
            ts2Dict = {}
            return(ts2Dict, errores)
        
        try:            
            for name in un_st:
                ts2Dict[name] = float(ts2Dict[name])
                
                if pd.isnull(ts2Dict[name]):
                   errores.append('Caracteres vacíos para ts2 min en parámetros')
                   ts2Dict = {}
                   return(ts2Dict, errores)                 
           
        except:
           errores.append('Error en procesamiento de hoja parámetros')
           ts2Dict = {}
           return(ts2Dict, errores)
           
        return(ts2Dict, errores)    
    
    except:
        errores.append('Error en parámetros')
        ts2Dict = {}
        return(ts2Dict, errores)    
    

        
        
def get_classPers(v, sheet_names, workbook):
    #leer la base y transformar a todos los ingresos en la data en objetos de la clase personas
    personas = []

    for names in sheet_names:
        df = workbook[names]
        pers_anio = []    
        i = 1  
        for index, row in df.iterrows():
            if names == sheet_names[0]:             
                pers_anio.append(Persona(row[2], row[3], row[4], row[5], row[6], row[7], row[18], i, i-1))            
            else:
                pers_anio.append(Persona(row[2], row[3], row[4], row[5], row[6], row[7], row[18], 'comp', i-1))        
            i+=1        
        personas.append(pers_anio)
        
        v.value += 1        
        
    return(personas)


def get_compPers(v, sheet_names, personas, ts2Dict):
    #Retornar la clase personas con las personas de los otros años asignadas su código único
    pers_comp = []

    pers_comp.append(personas[0])
    
    for i in range(len(sheet_names)-1):
        
        pos_piv = i      
        
        while int(float(sheet_names[pos_piv]) ) == int(float(sheet_names[i+1])):
            pos_piv -= 1
        
        pivot = pers_comp[pos_piv]
        toComp = personas[i+1]
        val = ts2Dict[str(int(float(sheet_names[pos_piv])))]
       
        
        std_toComp = sorted(toComp, key=op.attrgetter('nombre'))
#        gt1 = list(filter(lambda x: x.ts2 > val and x.cu != 'comp', pivot))
#        lt1 = list(filter(lambda x: x.ts2 <= val or x.cu == 'comp', pivot))
        gt1 = list(filter(lambda x: x.cu != 'comp', pivot))
        lt1 = list(filter(lambda x: x.cu == 'comp', pivot))
        
        
        for obj in gt1:
            pos = obj.comparador(std_toComp)
            if pos >= 0:
                if std_toComp[pos].cod_emp != obj.cod_emp:
                    obj.errores = obj.errores + 'Traspaso: De empresa ' + str(obj.cod_emp) + ' a empresa ' + str(std_toComp[pos].cod_emp) + ', '
                    
                if similar(std_toComp[pos].nombre, obj.nombre) < 0.78:
                    obj.errores = obj.errores + 'Revisar: El nombre presenta mucha diferencia, revisar si corresponde' + ', '
                
                if obj.ts2 <= val:
                    obj.errores = obj.errores + 'Revisar: Según ts2 no debería aparecer' + ', '
                
                ts2dif = obj.ts2 - std_toComp[pos].ts2
                if ts2dif < 0.974 or ts2dif > 1.05:
                    obj.errores = obj.errores + 'Revisar: El valor de ts2 puede tener un error' + ', '                    
                    
                std_toComp[pos].cu = obj.cu
            else:  
                if obj.ts2 > val:
                    obj.errores = obj.errores + 'Error: Debería aparecer en ' + sheet_names[i+1] + ', '
        
        pivot00 = sorted(gt1 + lt1, key=op.attrgetter('pos'))
        toComp11 = sorted(std_toComp, key=op.attrgetter('pos'))  
        
        pers_comp[pos_piv] = pivot00
        pers_comp.append(toComp11)
        
        v.value += 1
        
    
    return(pers_comp)  


def get_pdBase(v, sheet_names, workbook, pers_comp):
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
        v.value += 1    
    
    result = workbook[sheet_names[0]]
    for i in range(len(sheet_names)-1):
        data1 = workbook[sheet_names[i+1]]
        result = pd.merge(left=result, right=data1, on="CU", how="left")        
        v.value += 1
    
    return(result)     


def write_File(v, result, fileName):    
    #prepare database for presentation and write it to excel (or csv) file
    #result.to_csv("Result.csv")             
    book = load_workbook(filename='TEMPLATE.xlsm', read_only=False, keep_vba=True)
    v.value += 1         
    writer = pd.ExcelWriter(fileName[0:-5]+"_RESULT.xlsm", engine='openpyxl')     
    v.value += 1     
    writer.book = book    
    v.value += 1     
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)    
    v.value += 1 
    result.to_excel(writer, "Hoja1", header=True, index=False)    
    v.value += 5
    writer.save()  
    v.value += 5


def wrapper(v, v2, fileName, sheet_names, workbook, ts2Dict):  
    try:
        #self.text3.set('Paso 1 de 4: Cargando bases (objetos)')
        
        v2.value += 1 
        
        #print('Paso 1 de 4: Cargando bases (objetos)')
        personas = get_classPers(v, sheet_names, workbook)
        
        #self.text3.set('Paso 2 de 4: Comparando')
        
        v2.value += 1 
        
        #print('Paso 2 de 4: Comparando')            
        pers_comp = get_compPers(v, sheet_names, personas, ts2Dict) 
        
        
        #self.text3.set('Paso 3 de 4: Arreglando Base result.')
        v2.value += 1 
        #print('Paso 3 de 4: Arreglando Base result.')
        result = get_pdBase(v, sheet_names, workbook, pers_comp)
        
             
        #self.text3.set('Paso 4 de 4: Guardando archivo final')
        v2.value += 1 
        
        #print('Paso 4 de 4: Guardando archivo final')
        write_File(v, result, fileName)
        
        v2.value += 1 
        
    except Exception as e:
        print('Caught exception in wrapper function') 
        raise e
           
            
          























DELAY1 = 80


class Application(ttk.Frame):
    
    def __init__(self, parent):
        ttk.Frame.__init__(self, parent, name="frame")   
                 
        self.parent = parent
        self.initUI()
    
    def initUI(self):
        
        self.parent.title("Procesador Impuestos Diferidos (beta1)")
        
        self.logo = tk.PhotoImage(file='logo.gif') 
        self.label4 = tk.Label(self, image = self.logo)
        self.label4.place(x=190,y=5)          
        
        
        
        self.dataload_button = ttk.Button(
            self, text="Cargar Archivo", command=self.load_file)
        self.dataload_button.place(x=30, y=100)        
                
        self.run_button = ttk.Button(
            self, text="Ejecutar", command=self.run_prog)
        self.run_button.place(x=30, y=140)
        
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
        
  
        
        self.progressbar = ttk.Progressbar(self, orient="horizontal",
                                        length=200, mode="determinate")
        self.progressbar.place(x=30, y=180, width=200)       
        
        self.place(width=340, height=250)
        
        if getattr(sys, 'frozen', False):
            # frozen
            mdir = os.path.dirname(sys.executable)
        else:
            # unfrozen
            mdir = os.path.dirname(os.path.realpath(__file__))
        
        os.chdir(mdir)
        
        self.boolLoad = False

    def start_pb(self, d_length):
        self.progressbar["value"] = 0
        self.progressbar["maximum"] = d_length         
        
        
    def load_file(self):
        self.fileName = filedialog.askopenfilename(filetypes=[("xlsx files","*.xlsx")]) 
        try:   
            self.text1.set('Cargando Archivo.....')                          
            self.run_button.config(state=DISABLED)
            self.dataload_button.config(state=DISABLED)
            self.parent.update_idletasks() 
            

            pool = mp.Pool(mp.cpu_count())        
            self.workbook = pool.map(load_file, [self.fileName])[0]            
            self.sheet_names = list(self.workbook.keys())
            
            if 'Parametros' in self.sheet_names:                
                self.sheet_names.remove('Parametros')
                
                try:                
                    self.sheet_names.sort(key = float, reverse = True) 
                    
                except:
                    answer = messagebox.askokcancel("Procesador Imp. Diferidos", "La data cargada tiene problemas en el nombre de las hojas")
                    self.text1.set('[.]')
                    self.run_button.config(state=NORMAL)
                    self.dataload_button.config(state=NORMAL)
                    self.parent.update_idletasks()
                    return
                    
                
                self.text1.set('Validando Archivo.....') 
                self.parent.update_idletasks()
                self.datErr = []
                
                #Validate params
                self.ts2Dict, errores1 = validParam(self.workbook, self.sheet_names)
                self.datErr = self.datErr + errores1                 
                
                for sheet_name in self.sheet_names:
                    df = self.workbook[sheet_name]   
                    df, errores1 = validData(df, sheet_name)
                    self.workbook[sheet_name] = df
                    self.datErr = self.datErr + errores1                                    
                
                
                if len(self.datErr) > 0:
                   answer = messagebox.askokcancel("Procesador Imp. Diferidos", self.datErr)
                   with open(self.fileName[0:-5]+"_ERRORES.csv",'w') as resultFile:
                       wr = csv.writer(resultFile, dialect='excel')
                       for row in self.datErr:
                           wr.writerow([row])                   
                   
                   
                   self.text1.set('[.]')
                   self.run_button.config(state=NORMAL)
                   self.dataload_button.config(state=NORMAL)
                   self.parent.update_idletasks() 
                       
                else:          
                    self.totalIter = len(self.sheet_names)*4 + 12
                    self.start_pb(self.totalIter) 
                    self.boolLoad = True
                    self.text1.set('[Archivo Cargado]')
                    self.run_button.config(state=NORMAL)
                    self.dataload_button.config(state=NORMAL)
                    self.parent.update_idletasks() 
            else:
                answer = messagebox.askokcancel("Procesador Imp. Diferidos", "La data cargada no contiene la pestaña de nombre Parametros")
                self.text1.set('[.]')
                self.run_button.config(state=NORMAL)
                self.dataload_button.config(state=NORMAL)
                self.parent.update_idletasks() 
            
        except Exception as e:                
            answer = messagebox.askokcancel("Procesador Imp. Diferidos", e)
            if answer or not answer:
                self.parent.destroy()   
    
            
    def run_prog(self):
        
        if self.boolLoad:
        
            try:
                self.run_button.config(state=DISABLED)
                self.dataload_button.config(state=DISABLED)
            
                self.num = mp.Value('d', 0.0)
                self.num2 = mp.Value('d', 0.0)
            
                self.p1 = mp.Process(target=wrapper, args=(self.num, self.num2, self.fileName, self.sheet_names, self.workbook, self.ts2Dict))
            
                self.p1.start()
                self.after(DELAY1, self.onGetValue1)                 
            
            except Exception as e:
                answer = messagebox.askokcancel("Procesador Imp. Diferidos", e)
                if answer or not answer:
                    self.parent.destroy()   
        
        else:
            answer = messagebox.askokcancel("Procesador Imp. Diferidos", "No existe data correctamente cargada en el programa")
       
        
        
        
    def onGetValue1(self):
        
        if (self.p1.is_alive()):
            try:
                mydict = {0:'Iniciando...', 1:'Cargando Clase Personas', 2:'Comparando Bases', 3:'Generando Bases Resultantes', 4:'Escribiendo Resultados'}
                self.progressbar["value"] = self.num.value
                perc = '{:.2f}'.format(round(self.num.value/self.totalIter,4)*100)
                self.text2.set("%s%%" % perc) 
                self.text3.set('paso {} de 4: '.format(min(int(self.num2.value),4)) + mydict[min(int(self.num2.value),4)])
                self.after(DELAY1, self.onGetValue1)
                
                return
                
            except Exception as e:
                print(e)
                
        else:        
           try:
               if (self.num.value != self.totalIter) or (self.num2.value != 5):
                   answer = messagebox.askokcancel("Procesador Imp. Diferidos", "Error de procesamiento")
                   print((self.num.value,self.totalIter))
                   if answer or not answer:
                       self.parent.destroy()
               else:          
                   self.progressbar["value"] = self.totalIter
                   self.text2.set("%s%%" % 100)
                   self.run_button.config(state=NORMAL)
                   self.dataload_button.config(state=NORMAL)
                   answer = messagebox.askokcancel("Procesador Imp. Diferidos", "El proceso concluyó satisfactoriamente")
                   if answer or not answer:
                       self.parent.destroy()  
                   
           except Exception as e:
                print(e)
        
   
        
        
def main():
    root = Tk()
    root.geometry("400x250+400+150")
    app = Application(root)
    root.mainloop()        
        
        
if __name__ == '__main__':
    mp.freeze_support()
    main()         