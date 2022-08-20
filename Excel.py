import msvcrt
from re import L
import shutil
import sys
import time
import os
from datetime import datetime
import pyexcel
import xlwt
from win32com import client

class Excel_1:

    def __init__(self,cliente):
        self.cliente = cliente
        #Se crea el nombre el nombre de archivo temporal y su extensión
        self.file = self.cliente + ".xls"
    
    
    #Funcion para crear el archivo de excel
    def crear_xls(self,wb,dia,mes,ano):
        
        #Diccionario en el que se crea la cantidad de hojas que habrá en el archivo y sus columnas que habrá en cada una
        data = {'Datos': ["JUZGADO","RADICACION","DEMANDANTE","DEMANDANTE_CLIENTE",
                          "DEMANDADO","DEMANDADO_CLIENTE","CIUDAD","FECHA NOTIFICACION","ACTUACION","APODERADO"]}
        
        #Se agrega la hoja
        for key, nomHoja in enumerate(data):
            ws = wb.add_sheet(nomHoja)
            #Se agrega las columnas a cada hoja
            for clave, valor in enumerate(data[nomHoja]):
                ws.write(0, clave, valor)
        #Se guarda el archivo
        wb.save(f'.\\{ano}\\{mes}\\{dia}\\{self.file}')
        
    #Funcion para escribir en el excel creado anteriormente
    def escribir_xls(self,dato=None,dia=None,mes=None,ano=None):
        dato = list(dato)
        dato.pop(9)
        for i in range(len(dato)):
            try:
                dato[i] = dato[i].upper()
            except:
                pass
        #Se obtiene el archivo de excel para escribir
        wb = pyexcel.get_book(file_name=f'.\\{ano}\\{mes}\\{dia}\\{self.file}')
        #Se agregan los datos iniciales de el proceso
        wb.sheet_by_name('Datos').row += dato
        #Se guarda el archivo
        wb.save_as(f'.\\{ano}\\{mes}\\{dia}\\{self.file}')
    
    #Funcion en la que se guarda el archivo ya finalizado
    
    def convertir(self,dia,mes,ano):
            pyexcel.save_book_as(file_name=f'.\\{ano}\\{mes}\\{dia}\\{self.file}',
               dest_file_name=f'.\\{ano}\\{mes}\\{dia}\\{self.cliente}.xlsx')
            os.remove(f'.\\{ano}\\{mes}\\{dia}\\{self.file}')
    
    
    
    
    
    
    
    