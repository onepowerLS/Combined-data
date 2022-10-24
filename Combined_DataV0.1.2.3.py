
import sys
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook,workbook
import datetime as dt
from pathlib import Path  
import glob

"""
    Author: Motlatsi
    created: 17/Oct/2022

    This is the script that combines data of the given excel files into one master file designed to consolidate data for superclusters 
"""
class PopulateData:
    def __init__(self) -> None:
      PopulateData.instantiate_from_excel()
    
    #declaring lists to insert the sheets
    Pole_list=[]
    Network_list=[]
    DropLines_list=[]
    Connections_list=[]
    counter = 1

    #The function that removes empty spaces
    def remove_white_spaces(ws, row):
        for line in row:
            if line.value!=None and line.value !=' ': 
                return    
        ws.delete_rows(row[0].row, 1) # get the row number from the first cell, & remove the row
         
    @classmethod
    def instantiate_from_excel(cls):
        filename=''
        for file in files_in_cdir:
            if ('uGrid'in file and file.endswith('.xlsx'))  and sys.argv[1] not in file:
               filename=file
               wb = load_workbook(filename)
               print('\nPopulating....')

        for excel_file in files_in_odir:
           PopulateData.Pole_list.append(pd.read_excel(excel_file, index_col=0, sheet_name = 'PoleClasses'))
        ws= wb['PoleClasses']
        for excel in PopulateData.Pole_list:
            for row in dataframe_to_rows(excel,header = False): 
                print(excel)         
                ws.append(row) 
        for row in ws:# call a function to remove white spaces"
            PopulateData.remove_white_spaces(ws,row)

        for excel_file in files_in_odir:
            PopulateData.Network_list.append(pd.read_excel(excel_file, index_col=0, sheet_name = 'NetworkLength'))
        ws = wb['NetworkLength']
        for excel in PopulateData.Network_list:
            for row in dataframe_to_rows(excel,header = False):
                print(excel)    
                ws.append(row) 
        for row in ws:# call a function to remove white spaces
            PopulateData.remove_white_spaces(ws,row)

        for excel_file in files_in_odir:
            PopulateData.DropLines_list.append(pd.read_excel(excel_file,index_col=0, sheet_name = 'DropLines'))
        ws = wb['DropLines']
        for excel in PopulateData.DropLines_list:
            for row in dataframe_to_rows(excel,header = False): 
                print(excel)                   
                ws.append(row) 
        for row in ws:# call a function to remove white spaces
            PopulateData.remove_white_spaces(ws,row)

        for excel_file in files_in_odir:
            PopulateData.Connections_list.append(pd.read_excel(excel_file, index_col=0, sheet_name = 'Connections'))
        ws = wb['Connections']
        for excel in PopulateData.Connections_list:
            for row in dataframe_to_rows(excel,header = False): 
                print(excel)               
                ws.append(row) 
        for row in ws:# call a function to remove white spaces
            PopulateData.remove_white_spaces(ws,row)

        simdate = dt.datetime.today() # simulation date
        add0 = lambda x: '0'+str(x) if x < 10 else str(x) # Define a function that ensures double digits for day of the month
        filename = sys.argv[1]+'_'+str(simdate.year) + add0(simdate.month) + add0(simdate.day) + \
        '_' + add0(simdate.hour)+ add0(simdate.minute)+'_uGrid'+'.xlsx'

        wb.save(filename)
        print('\nDONE! Output File Path: ',path,'\\', filename )
        PopulateData.check_for_duplicates(filename)
        
    def check_for_duplicates(comb_file):
       print(f'\nChecking for duplicates....\n')
       df=pd.read_excel(comb_file,sheet_name='NetworkLength')
       duplicate_value = False
       for idx in range(len(df)):
           if(df.iat[idx,0]==df.iat[idx,1]):
               duplicate_value=True
               print('WARNING!!! Pole {} connects to the same Pole {} at row {} from PoleClasses in ({}).'.format(df.iat[idx,0],
                                                                                                                  df.iat[idx,1],
                                                                                                                  idx+2,comb_file)) 
       if not duplicate_value:
        print('No duplicate entries found!')
       
if __name__ == '__main__':
    try:
        if len(sys.argv)>1:
           SOURCE_DIR = sys.argv[1]
           files_in_odir = glob.glob(SOURCE_DIR + "/*.xlsx")
           path=os.getcwd()
           files_in_cdir = os.listdir(path)
    except:
        pass
    #class instantiation
    PopulateData()
    
