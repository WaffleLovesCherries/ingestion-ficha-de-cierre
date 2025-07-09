import io
import json
import logging
import os
import sys
import pandas as pd

sys.path.append(
    os.path.dirname(
        os.path.dirname(
            os.path.abspath(__file__)
        )
    )
)
from utils import *

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)

SHAREPOINT_ROOT_PATH = os.getenv('SHAREPOINT_ROOT_PATH')
SHAREPOINT_TARGET_FOLDERS = json.loads(os.getenv('SHAREPOINT_TARGET_FOLDERS', '[]'))
SHAREPOINT_OBSERVED_PATH = SHAREPOINT_ROOT_PATH + os.getenv('SHAREPOINT_OBSERVED_PATH')
SHAREPOINT_RECORDS_PATH = SHAREPOINT_ROOT_PATH + os.getenv('SHAREPOINT_RECORDS_PATH')

loader = SharepointLoader()
logging.info("Fetching metadata from SharePoint...")
files = loader.get_files(SHAREPOINT_ROOT_PATH, SHAREPOINT_TARGET_FOLDERS)
files = pd.DataFrame(files)

try: 
    logging.info("Loading previously observed files...")
    prev_loaded = loader.load_file(
        SHAREPOINT_OBSERVED_PATH, 
        as_format='dataframe'
    )
except Exception as e:
    logging.warning(f"Error loading file: {e}")
    prev_loaded = pd.DataFrame(columns=['UniqueId','Codigo','FichaCierre','Name','ServerRelativeUrl','TimeLastModified'])

logging.info("Preview of previously loaded files:\n%s", prev_loaded.head(3))

prev = prev_loaded[['UniqueId', 'TimeLastModified', 'Codigo', 'FichaCierre']].copy()
post = files[['UniqueId', 'TimeLastModified']].copy()
prev['present'], post['present'] = 1, 1

comparison = pd.merge(prev, post, on='UniqueId', how='outer', suffixes=(None, '_new'))
comparison[['present', 'present_new']] = comparison[['present', 'present_new']].fillna(0)

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
logging.info("Files to search for: %s", search)
logging.info("New files to add: (count: %d)", len(add))
logging.info("Files to modify: (count: %d)", len(modify))
logging.info("Unchanged files: (count: %d)", len(persist))

add, add_results = loader.process_wbs(add)
modify, modify_results = loader.process_wbs(modify)

try: 
    logging.info("Loading Banco de Fichas de Cierre...")
    persist_results = loader.load_file(
        SHAREPOINT_RECORDS_PATH, 
        as_format='dataframe'
    )
except Exception as e:
    logging.error(f"Error loading file: {e}")
    persist_results = pd.DataFrame(columns=['Codigo','Retos','AccionesDeMitigacion','LeccionesAprendidas'])

persist_results.merge(persist[['Codigo']], on='Codigo')

def replace_linebreaks(df):
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].str.replace('\r\n', '\\n', regex=False).str.replace('\n', '\\n', regex=False)
    return df

excel_prev_loaded_buffer = io.BytesIO()
pd.concat([persist, add, modify], axis=0).to_excel(excel_prev_loaded_buffer)
excel_prev_loaded_buffer.seek(0)

excel_fichas_buffer = io.BytesIO()
replace_linebreaks(pd.concat([persist_results, add_results, modify_results], axis=0)).to_excel(excel_fichas_buffer, index=False)
excel_fichas_buffer.seek(0)
logging.info("Saving updated Banco de Fichas de Cierre to SharePoint...")
loader.save_file(
    SHAREPOINT_RECORDS_PATH,
    excel_fichas_buffer
)

logging.info("Saving updated Archivos Observados to SharePoint...")
loader.save_file(
    SHAREPOINT_OBSERVED_PATH,
    excel_prev_loaded_buffer
)

