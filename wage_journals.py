# -*- coding: utf-8 -*-
"""
Created on Sun Apr 23 14:16:34 2023

@author: adrnna
"""
import xlwings as xw
import pandas as pd
import PyPDF2
import shutil 
import os
from tqdm import tqdm
import time


class ExcelExploded(Exception):
    pass

#______________________________________________________________________________

# # Get PIDs for all open instances of Excel
# pids = [xw.apps[k].pid for k in xw.apps.keys()]

# # Kill all Excel processes
# for pid in pids:
#     os.kill(pid, 9)
#______________________________________________________________________________

pdf_folder_path = input("Enter the path of the folder with the pdf files: ")
filter_pdfs = [pdf_name for pdf_name in os.listdir(pdf_folder_path) if '.pdf' in pdf_name]
filepath_list = [os.path.join(pdf_folder_path, pdf_name) for pdf_name in filter_pdfs]

xlsm_folder_path = input("Enter the path of the folder with the 'testing_vba.xlsm' file: ")
xlsm_src_name = 'testing_vba.xlsm'
xlsm_dst_name = 'testing_vba_copy.xlsm'

xlsm_src_file_path = os.path.join(xlsm_folder_path, xlsm_src_name)
xlsm_dst_file_path = os.path.join(xlsm_folder_path, xlsm_dst_name)

# =============================================================================
# FUNCTIONS
# =============================================================================

#check if the pdf has multiple pages, if so - split them 
def pdf_split(filepath, pdf_name):
    
    with open(filepath, 'rb') as file:
        # Create a PDF reader object
        pdf_reader = PyPDF2.PdfReader(file)
        
        numb_pages = len(pdf_reader.pages)
        
        if numb_pages > 1: 
            # Iterate through each page of the PDF file
            split_pdf_path_list = []
            split_pdf_name_list = []
            for page_num in range(len(pdf_reader.pages)):
                
                # Create a new PDF writer object
                pdf_writer = PyPDF2.PdfWriter()
        
                # Add the current page to the PDF writer object
                pdf_writer.add_page(pdf_reader.pages[page_num])
        
                # Write the current page to a new PDF file
                split_pdf_name = f'{os.path.splitext(pdf_name)[0]}_{page_num}.pdf'
                output_file_path = os.path.join(pdf_folder_path, split_pdf_name)
                with open(output_file_path, 'wb') as output_file:
                    pdf_writer.write(output_file)
                print('pdf split')
                split_pdf_path_list.append(output_file_path)
                split_pdf_name_list.append(split_pdf_name)
                
            return split_pdf_path_list, split_pdf_name_list
#______________________________________________________________________________      

def excel_macro(xlsm_src_file_path, xlsm_dst_file_path, filepath):
    try: 
        # copy the file
        shutil.copy(xlsm_src_file_path, xlsm_dst_file_path)
        
        # Create an invisible Excel application
        app = xw.App(visible=False)
        # while len(xw.apps.keys()) == 0 or not all(xw.apps[k].books for k in xw.apps.keys()):
        #     time.sleep(1)
        #time.sleep(10)
        
        # Create a new workbook
        wb = xw.Book('testing_vba_copy.xlsm')
        print('temporary excel file created')
        
        macro1 = wb.macro('Module1.import_pdf_ex')
        macro1(filepath)
        print('macro ran')
        
        wb.save()
        wb.close()
        
        # Read data from Excel file into pandas dataframe
        df_orig = pd.read_excel(xlsm_dst_file_path, engine = 'openpyxl', sheet_name='Sheet2', dtype=str)
        # drop the first two rows
        df_orig = df_orig.drop([0, 1])
        df_orig = df_orig.reset_index(drop=True)
        print('read data from excel')
        app.quit()
    except Exception as e:
        print(e)
        wb.close()
        app.quit()
        raise ExcelExploded ('Excel exploded :(')
    
    finally:
        # Quit the Excel application
        os.remove(xlsm_dst_file_path)
    
    return df_orig

# =============================================================================
# GET ALL DATA AND ORGANIZE IT
# =============================================================================

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
    check_content = df_orig[df_orig.iloc[:, 0].str.contains('Summe', na=False)].any().any()
    if check_content:
        # Find the cell with a delimiter value 
        row_num = df_orig[df_orig.iloc[:, 0].str.contains('Summe', na=False)].index[0]
        # Slice the DataFrame to keep only the rows above the row number
        df_check = df_orig.iloc[row_num:row_num+3, :].reset_index(drop=True)
        df_orig = df_orig.iloc[:row_num, :]
    
    new_cols_flat = redistrib_data(df_orig)
    
    return new_cols_flat, df_check

#______________________________________________________________________________

def concatenate_columns(column_set1, column_set2):
    concatenated_columns = []

    new_columns1 = []
    new_columns2 = []
    for cell in column_set1:
        if not isinstance(cell, str):
            new_columns1.append(cell)
        else:
            new_cell = cell.replace('.', '')
            new_cell = new_cell + '.'
            new_columns1.append(new_cell)
                    
    for cell in column_set2:
         if pd.notna(cell) and len(cell)==1:
             new_cell = '0' + cell
             new_columns2.append(new_cell)
         else:
             new_columns2.append(cell)        

    for i in range(len(new_columns1)):
        concatenated_columns.append(new_columns1[i] + new_columns2[i])

    return concatenated_columns
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

def clean_data(new_cols_flat):
    names1 = new_cols_flat[15]
    names2 = new_cols_flat[18]
    names3 = new_cols_flat[21]
    names_fixed = getnames(names1,names2, names3)
    
    # Call the function with the two column sets
    pers_nr = new_cols_flat[0]
    stkl = new_cols_flat[3]
    ANtyp = new_cols_flat[4]
    GV = new_cols_flat[5]
    Faktor = new_cols_flat[6]
    KiFrb = new_cols_flat[9]
    konf = new_cols_flat[12]
    freibetrag = new_cols_flat[7]
    StTg = new_cols_flat[13]
    BGRS = new_cols_flat[8]
    SVTg = new_cols_flat[14]
    steuerbrutto = concatenate_columns(new_cols_flat[16], new_cols_flat[19])
    PvB = concatenate_columns(new_cols_flat[17], new_cols_flat[20])
    lohnsteuer = concatenate_columns(new_cols_flat[22], new_cols_flat[25])
    pausch_lohnsteuer = concatenate_columns(new_cols_flat[23], new_cols_flat[26])
    kirsteuer = concatenate_columns(new_cols_flat[28], new_cols_flat[31])
    pau_kirsteuer = concatenate_columns(new_cols_flat[29], new_cols_flat[32])
    kindergeld = concatenate_columns(new_cols_flat[30], new_cols_flat[33])
    SolZ = concatenate_columns(new_cols_flat[34], new_cols_flat[37])
    pauschSolZ = concatenate_columns(new_cols_flat[35], new_cols_flat[38])
    kv_brutto = concatenate_columns(new_cols_flat[39], new_cols_flat[42])
    kv_beitrag_AN = concatenate_columns(new_cols_flat[40], new_cols_flat[43])
    kv_beitrag_AG = concatenate_columns(new_cols_flat[41], new_cols_flat[44])
    rv_brutto = concatenate_columns(new_cols_flat[45], new_cols_flat[48])
    rv_beitrag_AN = concatenate_columns(new_cols_flat[46], new_cols_flat[49])
    rv_beitrag_AG = concatenate_columns(new_cols_flat[47], new_cols_flat[50])
    av_brutto = concatenate_columns(new_cols_flat[54], new_cols_flat[57])
    av_beitrag_AN = concatenate_columns(new_cols_flat[55], new_cols_flat[58])
    av_beitrag_AG = concatenate_columns(new_cols_flat[56], new_cols_flat[59])
    pv_brutto = concatenate_columns(new_cols_flat[63], new_cols_flat[66])
    pv_beitrag_AN = concatenate_columns(new_cols_flat[64], new_cols_flat[67])
    pv_beitrag_AG = concatenate_columns(new_cols_flat[65], new_cols_flat[68])
    umlage1 = concatenate_columns(new_cols_flat[69], new_cols_flat[72])
    umlage2 = concatenate_columns(new_cols_flat[70], new_cols_flat[73])
    umlage_insolv = concatenate_columns(new_cols_flat[71], new_cols_flat[74])
    ges_brutto = concatenate_columns(new_cols_flat[75], new_cols_flat[78])
    nettobzg = concatenate_columns(new_cols_flat[76], new_cols_flat[79])
    auszbetrag = concatenate_columns(new_cols_flat[77], new_cols_flat[80])
    
    
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
    
    df_journal = df_journal.applymap(minus)
    print('cleaned data')
    return df_journal

#______________________________________________________________________________

def check_data(df_check, df_journal):
    steuerbrutto_check = concatenate_columns(df_check.iloc[1:,5], df_check.iloc[1:,6])
    lohnsteuer_check = concatenate_columns(df_check.iloc[1:,7], df_check.iloc[1:,8])
    kirchensteuer_check = concatenate_columns(df_check.iloc[1:,9], df_check.iloc[1:,10])
    kindergeld_check = concatenate_columns(df_check.iloc[:,11], df_check.iloc[:,12])
    kv_check = concatenate_columns(df_check.iloc[:,13], df_check.iloc[:,14])
    rv_check = concatenate_columns(df_check.iloc[:,15], df_check.iloc[:,16])
    av_check = concatenate_columns(df_check.iloc[:,18], df_check.iloc[:,19])
    pv_check = concatenate_columns(df_check.iloc[:,21], df_check.iloc[:,22])
    umlage_check = concatenate_columns(df_check.iloc[:,23], df_check.iloc[:,24])
    ges_brutto_check = concatenate_columns(df_check.iloc[:,25], df_check.iloc[:,26])
    
    all_checks = [steuerbrutto_check, lohnsteuer_check, kirchensteuer_check, kindergeld_check,
                  kv_check, rv_check, av_check, pv_check, umlage_check, ges_brutto_check]
    flat_list_checks = [round(float(item), 2) for sublist in all_checks for item in sublist]

    
    col_sums = []
    for col in ['Steuerbrutto', 'Pausch. verst. Bezüge', 'Lohnsteuer', 'Pausch. Lohnsteuer',
                'Kirchensteuer', 'Pausch. KiSt', 'Kindergeld', 'SolZ', 'Pausch. SolZ',
                'KV-Brutto', 'KV-Beitrag AN', 'KV-Beitrag AG', 'RV-Brutto', 'RV-Beitrag AN',
                'RV-Beitrag AG', 'AV-Brutto', 'AV-Beitrag AN', 'AV-Beitrag AG',
                'PV-Brutto', 'PV-Beitrag AN', 'PV-Beitrag AG', 'Umlage 1', 'Umlage 2',
                'Umlage Insolv.', 'Gesamtbrutto', 'Nettobezüge/-abzüge', 'Auszahlungsbetrag']:
        col_sum = df_journal[col].dropna().astype(float).sum()
        col_sum = round(col_sum, 2)
        col_sums.append(col_sum)
        
    #checking the sums   
    checked = (col_sums == flat_list_checks)
    
    return checked

# =============================================================================
# MAIN FUNCTION
# =============================================================================

def all_functions(filepath):
    df_orig = excel_macro(xlsm_src_file_path, xlsm_dst_file_path, filepath)
    new_cols_flat, df_check = organized_data(df_orig)
    df_journal = clean_data(new_cols_flat)    
    #checked = check_data(df_check, df_journal)
    return df_journal, df_check


def main(filepath, pdf_name):
    split_pdf_path_list = pdf_split(filepath, pdf_name)
    if split_pdf_path_list is not None:
        split_pdf_path_list, split_pdf_name_list = split_pdf_path_list
    
    #if the pdf has been split, run the functions on each page separately 
    if split_pdf_path_list is not None:
        df_journal_eachpage = {}
        for each_page, split_pdf_name in zip(split_pdf_path_list, split_pdf_name_list):
            filepath = each_page
            df_journal, df_check = all_functions(filepath)
            key_pdf = os.path.splitext(os.path.basename(split_pdf_name))[0]
            df_journal_eachpage[key_pdf] = df_journal
            
            # Get PIDs for all open instances of Excel
            pids = [xw.apps[k].pid for k in xw.apps.keys()]
            # Kill all Excel processes
            for pid in pids:
                os.kill(pid, 9)
                
        df_journal = pd.concat(df_journal_eachpage.values(), axis=0, ignore_index=True) 
        checked = check_data(df_check, df_journal)
        #remove the split files 
        for file in split_pdf_path_list:
            os.remove(file)
        return df_journal, checked
            
    else:
        df_journal, df_check = all_functions(filepath)
        checked = check_data(df_check, df_journal)
        return df_journal, checked

# =============================================================================
# RUN MAIN FUNCTION
# =============================================================================
#Initialize a dictionary to store the DataFrames
dfs_journals = {}
checked_list = {}

for filepath, pdf_name in tqdm(zip(filepath_list, filter_pdfs), total=len(filepath_list)):
    try:
        # Get PIDs for all open instances of Excel
        pids = [xw.apps[k].pid for k in xw.apps.keys()]
        # Kill all Excel processes
        for pid in pids:
            os.kill(pid, 9)
        
        df_journal, checked = main(filepath, pdf_name)
        dfs_journals[os.path.splitext(pdf_name)[0]] = df_journal
        checked_list[os.path.splitext(pdf_name)[0]] = checked
        print('waiting')
        time.sleep(10)
        
    except AttributeError as ae:
        print(f"1Document {pdf_name} failed. {ae}")
    except ExcelExploded as exc:
        print(f'2Document {pdf_name} failed. {exc}')    
    except Exception as e:
        print(f"3Document {pdf_name} failed. {e}")

#create xlsx file
concat_df = pd.concat(dfs_journals.values(), axis=0, ignore_index=True)        
concat_df.to_excel('wage_journals.xlsx', index=False)
