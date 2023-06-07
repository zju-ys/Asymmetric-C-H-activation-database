# -*- coding: utf-8 -*-
"""
Created on Wed Nov 30 08:14:51 2022

@author: silly
"""

import glob,os,re
import comtypes.client as w32
from rdkit import Chem
import pandas as pd
import xlrd,csv,glob
import numpy as np
def cdx2smi(file,ChemDraw):

    cdx_f = file
    tmp_compound = ChemDraw.Documents.Open(cdx_f)
    smiles = tmp_compound.Objects.Data("chemical/x-smiles")
    return smiles


def xlsx_to_csv(xlsx_f,csv_f):
    workbook = xlrd.open_workbook(xlsx_f)
    table = workbook.sheet_by_index(0)
    with open(csv_f, 'w', encoding='utf-8') as f:
        write = csv.writer(f)
        for row_num in range(table.nrows):
            row_value = table.row_values(row_num)
            write.writerow(row_value)

def get_name_list(name_keys):
    all_names = []
    for key in name_keys:
        names  = df[key].tolist()
        all_names.append(names)
    return all_names

def gen_smiles(name_list,type_):
    
    all_smiles = []
    
    for names in name_list:
        names_set = set(names)
        name_map = {}
        for r_n in names_set:
            if isinstance(r_n,str):
                cdx_f = '\\'.join(csv_f.split('\\')[:-1]) + '/%s/%s'%(type_,r_n)
                if os.path.exists(cdx_f + '.cdx'):
                    smi = cdx2smi(cdx_f + '.cdx',ChemDraw)
                elif os.path.exists(cdx_f + '.cdxml'):
                    smi = cdx2smi(cdx_f + '.cdxml',ChemDraw)
                else:
                    smi = ''
                    print('Not Existed',cdx_f)
                    not_existed_files.append('/'.join(cdx_f.split('/')[-2:]))
            else:
                smi = ''
            name_map[r_n] = smi
        smiles = [name_map[r_n] for r_n in names]
        all_smiles.append(smiles)
    smiles = [''.join(item) for item in np.array(all_smiles).T]
    return smiles
#%% csv 信息提取
ChemDraw = w32.CreateObject("ChemDraw.Application")

csv_files = glob.glob(r'C:\Users\ys\Asym_C-H_activation_database\literature\*\*.csv')
csv_files = [file for file in csv_files if not 'processed' in file and not 'not_existed_cdx_file' in file]
for csv_f in csv_files:
    print(csv_f)
    not_existed_files = []
    try:
        df = pd.read_csv(csv_f)
    except:
        df = pd.read_csv(csv_f,encoding='gbk')
    df_keys = list(df.keys())
    
    react_name_keys = [key for key in df_keys if 'Reactant' in key and 'Name' in key and '1' in key]
    react_name_keys_2 = [key for key in df_keys if 'Reactant' in key and 'Name' in key and '2' in key]
    prod_name_keys = [key for key in df_keys if 'Product' in key and 'Name' in key and '1' in key]
    prod_name_keys_2 = [key for key in df_keys if 'Product' in key and 'Name' in key and '2' in key]
    add_name_keys = [key for key in df_keys if 'Additive' in key and not 'mmount' in key and '1' in key and not 'step' in key]
    add_name_keys_2 = [key for key in df_keys if 'Additive' in key and not 'mmount' in key and '2' in key and not 'step' in key]
    add_name_keys_s2 = ['Additive 1 step 2']
    sol_name_keys = [key for key in df_keys if 'Solvent' in key and not 'Ratio' in key and not 'mmount' in key and '1' in key]
    sol_name_keys_2 = [key for key in df_keys if 'Solvent' in key and not 'Ratio' in key and not 'mmount' in key and '2' in key]
    cat_name_keys = [key for key in df_keys if '(Pre)Catalyst' in key and 'Name' in key and '1' in key]
    lig_name_keys = [key for key in df_keys if 'Ligand' in key and 'Name' in key and '1' in key]
    
    all_react_names = get_name_list(react_name_keys)
    all_react_names_2 = get_name_list(react_name_keys_2)
    all_prod_names = get_name_list(prod_name_keys)
    all_prod_names_2 = get_name_list(prod_name_keys_2)
    all_add_names = get_name_list(add_name_keys)
    all_add_names_2 = get_name_list(add_name_keys_2)
    all_add_names_s2 = get_name_list(add_name_keys_s2)
    all_sol_names = get_name_list(sol_name_keys)
    all_sol_names_2 = get_name_list(sol_name_keys_2)
    all_cat_names = get_name_list(cat_name_keys)
    all_lig_names = get_name_list(lig_name_keys)
    
    react_smiles = gen_smiles(all_react_names,'cdx')
    react_smiles_2 = gen_smiles(all_react_names_2,'cdx')
    prod_smiles = gen_smiles(all_prod_names,'cdx')
    prod_smiles_2 = gen_smiles(all_prod_names_2,'cdx')
    add_smiles = gen_smiles(all_add_names,'cdx')
    add_smiles_2 = gen_smiles(all_add_names_2,'cdx')
    add_smiles_s2 = gen_smiles(all_add_names_s2,'cdx')
    sol_smiles = gen_smiles(all_sol_names,'cdx')
    sol_smiles_2 = gen_smiles(all_sol_names_2,'cdx')
    cat_smiles = gen_smiles(all_cat_names,'catalyst cdx')
    lig_smiles = gen_smiles(all_lig_names,'catalyst cdx')
    
    
    key_inform = {'Reactant SMILES':react_smiles,'Reactant SMILES_2':react_smiles_2,'Product SMILES':prod_smiles,'Product SMILES_2':prod_smiles_2,
                  'Catalyst SMILES':cat_smiles,'Ligand SMILES':lig_smiles,
                  'Solvent SMILES':sol_smiles,'Solvent SMILES_2':sol_smiles_2,'Additive SMILES':add_smiles,'Additive SMILES_2':add_smiles_2,'Additive SMILES_s2':add_smiles_s2,}
    other_keys = set(df_keys) - set(react_name_keys) -set(react_name_keys_2) - set(prod_name_keys) - set(prod_name_keys_2) - set(add_name_keys) - set(add_name_keys_2)- set(add_name_keys_s2) -\
                 set(sol_name_keys) - set(sol_name_keys_2) - set(cat_name_keys) - set(lig_name_keys)
                 
    other_inform = {key:df[key].tolist() for key in other_keys}             
    all_inform = {}
    all_inform.update(key_inform)
    all_inform.update(other_inform)
    update_df = pd.DataFrame.from_dict(all_inform)
    
    update_df.to_csv(os.path.dirname(csv_f)+'/processed.csv')
    pd.DataFrame(not_existed_files).to_csv(os.path.dirname(csv_f)+'/not_existed_cdx_file.csv')
ChemDraw.quit()