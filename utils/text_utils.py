import re
import os
import json

from typing import Any, Union, List

REPLACEMENTS = json.loads(os.getenv('CHARACTER_REPLACEMENTS'))
VALID_CODE_PREFIXES = json.loads(os.getenv('VALID_CODE_PREFIXES'))

def to_bool(value: str) -> bool:
    """Convert string to boolean.
    Accepts: 'true', 'false', '1', '0', 'yes', 'no', 'on', 'off' (case insensitive)
    Parameters:
        value (str): The string to convert to boolean.
    Returns:
        out (bool): The converted boolean value.
    Raises:
        ValueError: If the string cannot be converted to a boolean.
    """
    if isinstance(value, bool): return value
    if value.lower() in ('true', '1', 'yes', 'on', 't', 'y'): return True
    elif value.lower() in ('false', '0', 'no', 'off', 'f', 'n'): return False
    else: raise ValueError(f"Cannot convert '{value}' to boolean")

def clean_str(value: str, allow_chars: Union[str, List[str]] = None) -> str:
    """Trim a string to remove unwanted characters.
    Parameters:
        value (str): The string to trim.
        allow_chars (str|list): Characters to keep in the string. If None, only alphanumeric characters are kept.
    Returns:
        out (str): The trimmed string.
    """
    if allow_chars is None:
        allow_chars = ''
    elif isinstance(allow_chars, str):
        allow_chars = list(allow_chars)
    
    pattern = f"[^a-záéíóúüñ0-9{''.join(allow_chars)}]"
    cleaned = re.sub(pattern, '', value.lower())

    for orig, repl in REPLACEMENTS.items():
        cleaned = cleaned.replace(orig, repl)
    return cleaned

def valid_code(code: Any) -> bool:
    """
    Checks if a code string is valid based on predefined prefixes and length.

    Parameters:
        code (Any): The code to validate.

    Returns:
        bool: True if the code is valid, False otherwise.
    """
    valid_prefixes = VALID_CODE_PREFIXES
    min_code_length = int(os.getenv('MIN_CODE_LENGTH', '4'))
    code_str = str(code).strip()
    return any(code_str.startswith(prefix) for prefix in valid_prefixes) and len(code_str) > min_code_length
