import os
import io
import json
import pandas as pd
import numpy as np
import openpyxl as px

from office365.sharepoint.files.file import File

from utils import *
loader = SharepointLoader()
files = loader.get_files('/sites/MicrositioProyectosFSD/Documentos compartidos/',['2. Gestión de Proyecto','3. Proyectos cerrados'])
files = pd.DataFrame(files)

try: 
    prev_loaded = loader.load_file(
        '/sites/MicrositioProyectosFSD/Documentos compartidos/6. Monitoreo/Fichas de Cierre/Archivos Observados.xlsx', 
        as_format='dataframe'
    )
except Exception as e:
    print(f"Error loading file: {e}")
    prev_loaded = pd.DataFrame(columns=['UniqueId','Codigo','FichaCierre','Name','ServerRelativeUrl','TimeLastModified'])

prev_loaded.head(3)

# Prepare previous and current file info
prev = prev_loaded[['UniqueId', 'TimeLastModified', 'Codigo', 'FichaCierre']].copy()
post = files[['UniqueId', 'TimeLastModified']].copy()
prev['present'], post['present'] = 1, 1

# Merge and compare
comparison = pd.merge(prev, post, on='UniqueId', how='outer', suffixes=(None, '_new'))
comparison[['present', 'present_new']] = comparison[['present', 'present_new']].fillna(0)

# Extract comparison results
search = list(
    comparison[comparison['present_new'] == 0][['UniqueId']]
    .merge(prev_loaded[['UniqueId','Codigo']], on='UniqueId', how='inner')
    ['Codigo']
)
global_search = list(prev_loaded['Codigo'])
add = (
    comparison[comparison['present'] == 0][['UniqueId']]
    .merge(files, on='UniqueId', how='inner')
    .set_index('UniqueId')
)
add.insert(0, 'Codigo', None)
add['FichaCierre'] = False
matches = (
    comparison[comparison['present'] == comparison['present_new']]
    [['UniqueId', 'TimeLastModified', 'TimeLastModified_new','Codigo', 'FichaCierre']]
)
modify = (
    matches
    .query('TimeLastModified < TimeLastModified_new')
    [['UniqueId','Codigo', 'FichaCierre']]
    .merge(files, on='UniqueId', how='inner')
    .set_index('UniqueId')
)
persist = (
    matches
    .query('TimeLastModified == TimeLastModified_new')
    [['UniqueId','Codigo', 'FichaCierre']]
    .merge(files, on='UniqueId', how='inner')
    .set_index('UniqueId')
)
print("Files to search for:", search)
print(f"\nNew files to add: (count: {len(add)})")
print(f"\nFiles to modify: (count: {len(modify)})")
print(f"\nUnchanged files: (count: {len(persist)})")

add, add_results = loader.process_wbs(add)
modify, modify_results = loader.process_wbs(modify)

try: 
    persist_results = loader.load_file(
        '/sites/MicrositioProyectosFSD/Documentos compartidos/6. Monitoreo/Fichas de Cierre/Banco de Fichas de Cierre.xlsx', 
        as_format='dataframe'
    )
except Exception as e:
    print(f"Error loading file: {e}")
    persist_results = pd.DataFrame(columns=['Codigo','Retos','AccionesDeMitigacion','LeccionesAprendidas'])

persist_results.merge(persist[['Codigo']], on='Codigo')

def replace_linebreaks(df):
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].str.replace('\r\n', '\\n', regex=False).str.replace('\n', '\\n', regex=False)
    return df

import io

# Save using loader to SharePoint
excel_prev_loaded_buffer = io.BytesIO()
pd.concat([persist, add, modify], axis=0).to_excel(excel_prev_loaded_buffer)
excel_prev_loaded_buffer.seek(0)

loader.save_file(
    '/sites/MicrositioProyectosFSD/Documentos compartidos/6. Monitoreo/Fichas de Cierre/Archivos Observados.xlsx',
    excel_prev_loaded_buffer
)

excel_fichas_buffer = io.BytesIO()
replace_linebreaks(pd.concat([persist_results, add_results, modify_results], axis=0)).to_excel(excel_fichas_buffer, index=False)
excel_fichas_buffer.seek(0)
loader.save_file(
    '/sites/MicrositioProyectosFSD/Documentos compartidos/6. Monitoreo/Fichas de Cierre/Banco de Fichas de Cierre.xlsx',
    excel_fichas_buffer
)
