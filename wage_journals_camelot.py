# -*- coding: utf-8 -*-
"""
Created on Sun Apr 23 14:16:34 2023

@author: adrnna
"""
import pandas as pd
import os
from tqdm import tqdm
import re
import camelot
import numpy as np
import json


pdf_folder_path = input("Enter the path of the folder with the pdf files: ")
#pdf_folder_path = 'path//to//folder'
filter_pdfs = [pdf_name for pdf_name in os.listdir(pdf_folder_path) if '.pdf' in pdf_name]
filepath_list = [os.path.join(pdf_folder_path, pdf_name) for pdf_name in filter_pdfs]

# =============================================================================
# FUNCTIONS: GET DATA WITH CAMELOT 
# =============================================================================

def camelot_reader(filepath):    
    #two dataframes with different settings combined to get easy to clean data
    table1 = camelot.read_pdf(filepath, pages='all', flavor='stream', split_text=False, 
                          edge_tol=10,
                          layout_kwargs={'detect_vertical': False,
                                         'line_overlap': 0.06,
                                         }) 
    
    table2 = camelot.read_pdf(filepath, pages='all', flavor='stream', split_text=False, 
                              edge_tol=10,
                              layout_kwargs={'detect_vertical': True,
                                             'line_overlap': 0.05,
                                             'char_margin': 0.077,
                                             }) 
    
    # combined the two extracted dataframes
    df_came1 = []
    df_came2 = []        
    for table in table1:
        df_came1_page = table.df.iloc[:, :2]   
        df_came1.append(df_came1_page)             
    for table in table2:
        df_came2_page = table.df.iloc[:, 1:]   
        df_came2.append(df_came2_page)       
                              
    df_orig = []
    for df1, df2 in zip(df_came1, df_came2):
        df_comb = pd.concat([df1, df2], axis=1, ignore_index=True)
        df_orig.append(df_comb)
    
    return df_orig

# =============================================================================
# FUNCTIONS: ORGANIZE AND CLEAN DATA
# =============================================================================

# reorganizes the data by grouping it every third value
def redistrib_data(df):
    data_to_redistribute_list = []
    new_cols_list = [] 
    for col in df.columns:
        data_to_redistribute = df[col]
        new_cols = [[data_to_redistribute[j] for j in range(k, len(data_to_redistribute), 3)] for k in range(3)]
        data_to_redistribute_list.append(list(data_to_redistribute))
        new_cols_list.append(new_cols)
    new_df = [item for sublist in new_cols_list for item in sublist]
    return new_df


def organized_data(df_orig):
    # Find the cell with a delimiter value: first number in the first column
    for idx, val in enumerate(df_orig.iloc[:, 0]):
        match = re.search(r'\d+', val)
        if match:
            first_num_row = idx
            break
    if first_num_row:
        df_orig = df_orig.iloc[first_num_row:, :].reset_index(drop=True)
    
    check_content = df_orig[df_orig.iloc[:, 0].str.contains('Summe', na=False)].any().any()
    if check_content:
        # Find the cell with a delimiter value 
        row_num = df_orig[df_orig.iloc[:, 0].str.contains('Summe', na=False)].index[0]
        # Slice the DataFrame to keep only the rows above the row number
        df_check = df_orig.iloc[row_num:row_num+3, :].reset_index(drop=True)
        df_orig = df_orig.iloc[:row_num, :]
        #go to the minus function
        df_orig = df_orig.applymap(minus)
    new_cols_flat = redistrib_data(df_orig)
    
    return new_cols_flat, df_check

#______________________________________________________________________________

def decimals(column_set):
    converted_values = []
    for value in column_set:
        #remove any periods or letters
        no_period = ''.join(filter(lambda char: char.isdigit() or char == '-', value))
        if not no_period:
           converted_values.append(no_period)
        else:
            converted_value = no_period[:-2] + '.' + no_period[-2:]
            if converted_value.startswith('.'):
                converted_value = '0' + converted_value
            converted_values.append(converted_value)

    return converted_values

#______________________________________________________________________________

def minus(val):
    if pd.notna(val) and val.endswith('-'):
        val = '-' + val[:-1]
        
    return val

#______________________________________________________________________________

def getnames(names1,names2, names3):
    names_conc = []
    # Loop through both lists and concatenate the elements at the same index
    for i in range(len(names1)):
        names_conc.append(names1[i] + names2[i] + names3[i])
        
    names_fixed = []
    for name in names_conc:
        if ',' in name: 
            comma_ind = name.index(',')
            if name[comma_ind+1] != ' ':
                new_name = name[:comma_ind+1] + ' ' + name[comma_ind+1:]
                names_fixed.append(new_name)
            else:
                names_fixed.append(name)
        else:
            names_fixed.append(name)
            
    return names_fixed

#______________________________________________________________________________

def split_by_type(original_list):
    list_dig = []
    list_str = []

    for item in original_list:
        if any(char.isdigit() for char in item):
            num = ''.join(filter(lambda char: char.isdigit() or char in ['.', ','], item))
            list_dig.append(num)
            list_str.append(''.join(filter(str.isalpha, item)))
        else:
            list_str.append(item)
            list_dig.append('')
            
    return (list_dig, list_str)

#______________________________________________________________________________

def clean_data(new_cols_flat):
    names1 = new_cols_flat[15]
    names2 = new_cols_flat[18]
    names3 = new_cols_flat[21]
    names_fixed = getnames(names1,names2, names3)
    
    pers_nr = new_cols_flat[0]
    stkl = new_cols_flat[3]
    ANtyp = new_cols_flat[4]
    GV = new_cols_flat[5]
    Faktor = new_cols_flat[6]
    KiFrb, konf = split_by_type(new_cols_flat[12])
    freibetrag = new_cols_flat[7]
    StTg = new_cols_flat[16]
    BGRS = new_cols_flat[8]
    SVTg = new_cols_flat[17]
    steuerbrutto = decimals(new_cols_flat[19])
    PvB = decimals(new_cols_flat[20])
    lohnsteuer = decimals(new_cols_flat[22])
    pausch_lohnsteuer = decimals(new_cols_flat[23])
    kirsteuer = decimals(new_cols_flat[25])
    pau_kirsteuer = decimals(new_cols_flat[26])
    kindergeld = decimals(new_cols_flat[27])
    SolZ = decimals(new_cols_flat[28])
    pauschSolZ = decimals(new_cols_flat[29])
    kv_brutto = decimals(new_cols_flat[30])
    kv_beitrag_AN = decimals(new_cols_flat[31])
    kv_beitrag_AG = decimals(new_cols_flat[32])
    rv_brutto = decimals(new_cols_flat[36])
    rv_beitrag_AN = decimals(new_cols_flat[37])
    rv_beitrag_AG = decimals(new_cols_flat[38])
    av_brutto = decimals(new_cols_flat[42])
    av_beitrag_AN = decimals(new_cols_flat[43])
    av_beitrag_AG = decimals(new_cols_flat[44])
    pv_brutto = decimals(new_cols_flat[45])
    pv_beitrag_AN = decimals(new_cols_flat[46])
    pv_beitrag_AG = decimals(new_cols_flat[47])
    umlage1 = decimals(new_cols_flat[48])
    umlage2 = decimals(new_cols_flat[49])
    umlage_insolv = decimals(new_cols_flat[50])
    ges_brutto = decimals(new_cols_flat[51])
    nettobzg = decimals(new_cols_flat[52])
    auszbetrag = decimals(new_cols_flat[53])
    
    
    df_journal = pd.DataFrame({'Pers.-Nr.': pers_nr, 'St. Kl.': stkl,
                                'AN-Typ': ANtyp, 'GV': GV, 'Faktor': Faktor,
                                'Ki.Frb.': KiFrb, 'Konf. AN/Eheg.': konf, 
                                'Freibetrag': freibetrag, 'St.Tg.': StTg,
                                'BGRS': BGRS, 'SV.Tg.': SVTg, 'Name': names_fixed,
                                'Steuerbrutto': steuerbrutto, 'Pausch. verst. Bezüge': PvB,
                                'Lohnsteuer': lohnsteuer, 'Pausch. Lohnsteuer': pausch_lohnsteuer,
                                'Kirchensteuer': kirsteuer, 'Pausch. KiSt': pau_kirsteuer,
                                'Kindergeld': kindergeld, 'SolZ': SolZ, 'Pausch. SolZ': pauschSolZ,
                                'KV-Brutto': kv_brutto, 'KV-Beitrag AN': kv_beitrag_AN, 'KV-Beitrag AG': kv_beitrag_AG,
                                'RV-Brutto': rv_brutto, 'RV-Beitrag AN': rv_beitrag_AN, 'RV-Beitrag AG': rv_beitrag_AG,
                                'AV-Brutto': av_brutto, 'AV-Beitrag AN': av_beitrag_AN, 'AV-Beitrag AG': av_beitrag_AG,
                                'PV-Brutto': pv_brutto, 'PV-Beitrag AN': pv_beitrag_AN, 'PV-Beitrag AG': pv_beitrag_AG,
                                'Umlage 1': umlage1, 'Umlage 2': umlage2, 'Umlage Insolv.': umlage_insolv,
                                'Gesamtbrutto': ges_brutto, 'Nettobezüge/-abzüge': nettobzg, 'Auszahlungsbetrag': auszbetrag})
    
    return df_journal

# =============================================================================
# FUNCTIONS: CHECK DATA
# =============================================================================

def check_data(df_check, df_journal):
    steuerbrutto_check = decimals([df_check.iloc[1,6]])
    PvB_check = decimals([df_check.iloc[2,6]])
    lohnsteuer_check = decimals([df_check.iloc[1,7]])
    pausch_lohnsteuer_check = decimals([df_check.iloc[2,7]])
    kirchensteuer_check = decimals([df_check.iloc[1,8]])
    pau_kirsteuer_check = decimals([df_check.iloc[2,8]])
    kindergeld_check = decimals([df_check.iloc[0,9]])
    SolZ_check = decimals([df_check.iloc[1,9]])
    pauschSolZ_check = decimals([df_check.iloc[2,9]])
    kv_check = decimals([df_check.iloc[0,10]])
    kv_beitrag_AN_check = decimals([df_check.iloc[1,10]])
    kv_beitrag_AG_check = decimals([df_check.iloc[2,10]])
    rv_check = decimals([df_check.iloc[0,12]])
    rv_beitrag_AN_check = decimals([df_check.iloc[1,12]])
    rv_beitrag_AG_check = decimals([df_check.iloc[2,12]])
    av_check = decimals([df_check.iloc[0,14]])
    av_beitrag_AN_check = decimals([df_check.iloc[1,14]])
    av_beitrag_AG_check = decimals([df_check.iloc[2,14]])
    pv_check = decimals([df_check.iloc[0,15]])
    pv_beitrag_AN_check = decimals([df_check.iloc[1,15]])
    pv_beitrag_AG_check = decimals([df_check.iloc[2,15]])
    umlage1_check = decimals([df_check.iloc[0,16]])
    umlage2_check = decimals([df_check.iloc[1,16]])
    umlage_insolv_check = decimals([df_check.iloc[2,16]])
    ges_brutto_check = decimals([df_check.iloc[0,17]])
    nettobzg_check = decimals([df_check.iloc[1,17]])
    auszbetrag_check = decimals([df_check.iloc[2,17]])
    
    all_checks = [steuerbrutto_check, PvB_check, lohnsteuer_check, pausch_lohnsteuer_check,
                  kirchensteuer_check, pau_kirsteuer_check, kindergeld_check, SolZ_check,
                  pauschSolZ_check, kv_check, kv_beitrag_AN_check, kv_beitrag_AG_check,
                  rv_check, rv_beitrag_AN_check, rv_beitrag_AG_check,
                  av_check, av_beitrag_AN_check, av_beitrag_AG_check,
                  pv_check, pv_beitrag_AN_check, pv_beitrag_AG_check,
                  umlage1_check, umlage2_check, umlage_insolv_check,
                  ges_brutto_check, nettobzg_check, auszbetrag_check]
    flat_list_checks = [round(float(item), 2) for sublist in all_checks for item in sublist]


    col_sums = []
    for col in ['Steuerbrutto', 'Pausch. verst. Bezüge', 'Lohnsteuer', 'Pausch. Lohnsteuer',
                'Kirchensteuer', 'Pausch. KiSt', 'Kindergeld', 'SolZ', 'Pausch. SolZ',
                'KV-Brutto', 'KV-Beitrag AN', 'KV-Beitrag AG', 'RV-Brutto', 'RV-Beitrag AN',
                'RV-Beitrag AG', 'AV-Brutto', 'AV-Beitrag AN', 'AV-Beitrag AG',
                'PV-Brutto', 'PV-Beitrag AN', 'PV-Beitrag AG', 'Umlage 1', 'Umlage 2',
                'Umlage Insolv.', 'Gesamtbrutto', 'Nettobezüge/-abzüge', 'Auszahlungsbetrag']:
        df_journal = df_journal.replace('', np.nan)
        col_sum = df_journal[col].dropna().astype(float).sum()
        col_sum = round(col_sum, 2)
        col_sums.append(col_sum)
        
    #checking the sums   
    checked = (col_sums == flat_list_checks)
    
    return checked

# =============================================================================
# MAIN 
# =============================================================================

def all_functions(filepath):
    df_orig = camelot_reader(filepath)
    # loop for each page
    df_journal_all = []
    for each_df in df_orig:
        new_cols_flat, df_check = organized_data(each_df)
        df_journal = clean_data(new_cols_flat)    
        df_journal_all.append(df_journal)
    concat_journal = pd.concat(df_journal_all, axis=0)
    checked = check_data(df_check, concat_journal)
    return concat_journal, checked


def main(filepath, pdf_name):
    df_journal, checked = all_functions(filepath)
    return df_journal, checked

# =============================================================================
# RUN MAIN 
# =============================================================================

#Initialize a dictionary to store the DataFrames
dfs_journals = {}
#Store booleans indicating whether the summed values match with the sums in the documents
checked_list = {}

for filepath, pdf_name in tqdm(zip(filepath_list, filter_pdfs), total=len(filepath_list)):
    try:        
        df_journal, checked = main(filepath, pdf_name)
        dfs_journals[os.path.splitext(pdf_name)[0]] = df_journal
        checked_list[os.path.splitext(pdf_name)[0]] = checked
        
    except Exception as e:
        print(f"3Document {pdf_name} failed. {e}")

# export journal to xlsx file
concat_df = pd.concat(dfs_journals.values(), axis=0, ignore_index=True)        
concat_df.to_excel('wage_journals.xlsx', index=False)

# export checked_list to a txt file
with open('wage_data_check.txt', 'w') as f:
    for key, value in checked_list.items():
         f.write(key + ': ' + str(value) + '\n')