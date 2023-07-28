from PyPDF2 import PdfReader
from openpyxl import Workbook
import re
import os
from natsort import natsorted


def leer_pagina(numero):
    pagina = reader.pages[numero]
    texto = pagina.extract_text(0)
    return texto.splitlines()

def generar_excel(entradas, salidas, excelpath, variablepath):
    excel = Workbook()
    hoja_entradas = excel.active
    hoja_entradas.title = 'PLC Tags'
    hoja_entradas['A1'].value = "Name"
    hoja_entradas['B1'].value = "Path"
    hoja_entradas['C1'].value = "Data Type"
    hoja_entradas['D1'].value = "Logical Address"
    hoja_entradas['E1'].value = "Comment"
    hoja_entradas['F1'].value = "Hmi Visible"
    hoja_entradas['G1'].value = "Hmi Accessible"
    hoja_entradas['H1'].value = "Hmi Writeable"
    hoja_entradas['I1'].value = "Typeobject ID"
    hoja_entradas['J1'].value = "Version ID"

    entradas = natsorted(entradas)
    salidas = natsorted(salidas)

    aux = 0
    for idx, entrada in enumerate(entradas, start=2):
        separador = re.split('(E\d\.\d\s|E\d\d\.\d\s|E\d\d\d\.\d\s|E\d\d\d\d\.\d\s)', entrada)
        hoja_entradas[f'D{idx}'].value = separador[1].replace("E","%I")
        hoja_entradas[f'A{idx}'].value = separador[2].replace(" ","_").rsplit("_", 1)[0]
        hoja_entradas[f'B{idx}'].value = variablepath
        hoja_entradas[f'C{idx}'].value = "Bool"
        hoja_entradas[f'F{idx}'].value = "True"
        hoja_entradas[f'G{idx}'].value = "True"
        hoja_entradas[f'H{idx}'].value = "True"
        aux = idx + 1

    for idx, salida in enumerate(salidas, start=aux):
        separador = re.split('(A\d\.\d\s|A\d\d\.\d\s|A\d\d\d\.\d\s|A\d\d\d\d\.\d\s)', salida)
        hoja_entradas[f'D{idx}'].value = separador[1].replace("A","%Q")
        hoja_entradas[f'A{idx}'].value = separador[2].replace(" ","_").rsplit("_", 1)[0]
        hoja_entradas[f'B{idx}'].value = variablepath
        hoja_entradas[f'C{idx}'].value = "Bool"
        hoja_entradas[f'F{idx}'].value = "True"
        hoja_entradas[f'G{idx}'].value = "True"
        hoja_entradas[f'H{idx}'].value = "True"
    
    excel.save(excelpath)

def generar_listados(rangopaginas):
    entradas = set()
    salidas = set()

    e1 = re.compile(r"E\d\.\d\s[A-Za-z]|E\d\d\.\d\s[A-Za-z]|E\d\d\d\.\d\s[A-Za-z]|E\d\d\d\d\.\d\s[A-Za-z]", re.IGNORECASE)
    e2 = re.compile(r"DI2E\d\.\d\s[A-Za-z]|DI2E\d\d\.\d\s[A-Za-z]|DI2E\d\d\d\.\d\s[A-Za-z]|DI2E\d\d\d\d\.\d\s[A-Za-z]", re.IGNORECASE)
    a1 = re.compile(r"A\d\.\d\s[A-Za-z]|A\d\d\.\d\s[A-Za-z]|A\d\d\d\.\d\s[A-Za-z]|A\d\d\d\d\.\d\s[A-Za-z]", re.IGNORECASE)
    a2 = re.compile(r"MA\d\.\d\s[A-Za-z]|MA\d\d\.\d\s[A-Za-z]|MA\d\d\d\.\d\s[A-Za-z]|MA\d\d\d\d\.\d\s[A-Za-z]", re.IGNORECASE)
    r5 = re.compile(r"Reserve", re.IGNORECASE)

    for p in rangopaginas:
        lineas = leer_pagina(p)
        for linea in lineas:
            if (re.match(e1, linea) or re.match(e2, linea)) and not re.search(r5, linea):               
                entradas.add(linea.replace("DI2","").replace("+","").replace("ü","ue").replace("ö","oe").replace("ä","ae").rsplit("-", 1)[0])
            elif (re.match(a1, linea) or re.match(a2, linea)) and not re.search(r5, linea):
                salidas.add(linea.replace("M","").replace("+","").replace("ü","ue").replace("ö","oe").replace("ä","ae").rsplit("-", 1)[0])

    return list(entradas), list(salidas)

if __name__ == "__main__":

    escritorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    excel_path = os.path.join(escritorio,'Variables.xlsx')

    print('Introducir ruta del archivo (EJ: C:\\Users\\User\\Documents\\archivo.pdf):')
    pdf_path = input()
    reader = PdfReader(pdf_path)

    print("¿Desde que página quieres buscar?: ")
    lowpage = int(input())
    print("¿Hasta que página quieres buscar?: ")
    highpage = int(input())
    rangopaginas = range(lowpage, highpage)
    print("Nombre para la tabla de variables: ")
    tablavariables = input()

    print("Buscando señales...")
    entradas, salidas = generar_listados(rangopaginas)
    os.system('cls')
    print("Generando Excel...")
    generar_excel(entradas,salidas, excel_path, tablavariables)
    os.system('cls')
    print("Terminado")
    input()
