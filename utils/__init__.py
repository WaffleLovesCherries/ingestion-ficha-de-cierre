# Import all components to make them available when importing the utils package

from .SharepointLoader import SharepointLoader
from .text_utils import to_bool, clean_str, valid_code
from .env_utils import load_environment_variables
from .excel_utils import setter, wbs_check

# This is what will be available when you do `from utils import *`
__all__ = [
    'SharepointLoader',
    'to_bool', 
    'clean_str',
    'valid_code',
    'load_environment_variables',
    'setter',
    'wbs_check'
]

# Entry point for testing
if __name__ == "__main__":
    loader = SharepointLoader()
    files = loader.get_files('/sites/MicrositioProyectosFSD/Documentos compartidos/', ['2. Gestión de Proyecto', '3. Proyectos cerrados'])
    print(len(files))