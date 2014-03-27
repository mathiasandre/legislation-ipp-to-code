# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd
import datetime

# Architecture : 
# un xlsx contient des sheets qui contiennent des variables, chaque sheet ayant un vecteur de dates/

def import_xls(param_name):# Travail sur Prestations uniquement pour l'instant
    return pd.ExcelFile('C:/Users/m.guillot/My Documents/Aptana Studio 3 Workspace/Param_IPP/Baremes IPP - '+param_name+'.xlsx')

def clean_sheet(sheet_name):
    ''' Cleaning excel sheets and creating small database'''
    sheet = xlsxfile.parse(sheet_name, index_col = None)
    #print sheet.columns.values # Column names
    print sheet.columns.values
    for col in sheet.columns.values:     #Conserver les bonnes colonnes : on drop tous les "Unnamed"
        if (col[0:7] == 'Unnamed'):
            sheet = sheet.drop([col],axis = 1)

    #Conserver les bonnes lignes : on drop une ou deux lignes du début
    #TODO: utiliser cette ligne pour faire des labels/descriptions
    list_prestation = ['PAJE_CM', 'ALF1', 'ALF4', 'ALF6', 'ALF8', 'ALF9', 'ALF12'] 
    if sheet_name in list_prestation: #TODO: eventuellement améliorer en testant si la cellule d' est un nombre et delete sinon
        sheet = sheet.ix[2:]  
    else: 
        sheet = sheet.ix[1:]
    #TODO: droper les lignes de la fin quand il n'y a pas de date mais des trucs écrits

    # Handle dates

    # Handle currencies
    def _francs_to_euros(var):
        ''' Fonction qui converti les francs en euro'''
        sheet[var][pd.DatetimeIndex(sheet['date']).year <= 2001] = np.divide(sheet[var][pd.DatetimeIndex(sheet['date']).year <= 2001],6.55957) # tout est en euro du coup
        return sheet[var]
    #TODO: utiliser la fonction uniquement quand on a un montant nominal (mais pas pour les taux ou les conditions d'ages par ex.) 
        # => comment savoir s'il est nominal ou non ? (idée : tester le résultat est proche de l'année d'après (par exemple, maximum 10% de variation ?)
        # sinon, on fait la liste à la main
    return sheet
    
    
if __name__ == '__main__':

    xlsxfile =   import_xls('Prestations')
    sheet_names = xlsxfile.sheet_names # get all sheet names

    bmaf = clean_sheet('BMAF')


