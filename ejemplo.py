# This Python file uses the following encoding: utf-8
import sys
import os
import xlwings as xw

def main():
    xw.sheets('Hoja1').activate()
    xw.Range('A1').value = u'Hola! Ya estás usando Python!'
    xw.Range('A2').value = u'El directorio de trabajo actual es:'
    xw.Range('A3').value = os.getcwd()
