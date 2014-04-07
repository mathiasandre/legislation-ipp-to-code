# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd
import math
import os
# Architecture : 
# un xlsx contient des sheets qui contiennent des variables, chaque sheet ayant un vecteur de dates

def import_xls(param_name):
    path = os.path.dirname(__file__)
    return pd.ExcelFile(path+'/Baremes IPP - '+param_name+'.xlsx')
    
def clean_sheet(sheet_name):
    ''' Cleaning excel sheets and creating small database'''

    sheet = xlsxfile.parse(sheet_name, index_col = None)

    # Conserver les bonnes colonnes : on drop tous les "Unnamed"
    for col in sheet.columns.values:     
        if (col[0:7] == 'Unnamed'):
            sheet = sheet.drop([col],axis = 1)
   
    def _is_var_nan(row,col):
        fusion = False
        if isinstance(sheet.iloc[row,col],float):
            fusion = math.isnan(sheet.iloc[row,col])
        return fusion
    
    # Conserver les bonnes lignes : on drop s'il y a du texte ou du NaN dans la colonne des dates
    sheet['date_renseignees'] = False
    for i in range(0,sheet.shape[0]):
        sheet.loc[i,['date_renseignees']] = isinstance(sheet.iat[i,0],unicode) | _is_var_nan(i,0)
    sheet = sheet[sheet.date_renseignees == False]
    sheet = sheet.drop(['date_renseignees'],axis = 1)
    
    # S'il y a du texte au milieu du tableau (explications par exemple) => on le transforme en NaN
    for col in range(0,sheet.shape[1]):
        for row in range(0,sheet.shape[0]):
            if isinstance(sheet.iloc[row,col],unicode):
                sheet.iat[row,col] = 'NaN'
    # TODO: Handle currencies (Pb : on veut ne veut diviser que les montants et valeurs monétaires mais pas les taux ou conditions).
    # TODO: Utiliser les lignes supprimées du début pour en faire des lables
    # TODO: Utiliser les lignes supprimées de la fin et de la droite donner des informations sur la législation (références, notes...)
    return sheet
    
    
if __name__ == '__main__':

    xlsxfile =   import_xls('Prestations')
    sheet_names = xlsxfile.sheet_names
#    list_prestation = ['PAJE_CM', 'ALF1', 'ALF4', 'ALF6', 'ALF8', 'ALF9', 'ALF12'] La liste des cas compliqués pour Prestation
    sheet_names = [ v for v in sheet_names if not v.startswith('Sommaire')|v.startswith('Outline') | v.startswith('Barème IGR')]
    print sheet_names
    dic={}
    for sheet in  sheet_names:
        dic[sheet]= clean_sheet(sheet)
