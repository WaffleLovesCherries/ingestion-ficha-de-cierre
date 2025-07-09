import os

from dotenv import load_dotenv

def load_environment_variables():
    """Load environment variables from *.env* file, and optionally from a secret path if `BIG_BROTHER_WATCHING` is set.
    This is useful for loading sensitive information like API keys or database credentials by protecting them from GitHub Copilot.
    Raises:
        ValueError: If the string cannot be converted to a boolean.
    """
    load_dotenv(override=True)
    if os.getenv('BIG_BROTHER_WATCHING'):
        load_dotenv(os.getenv('SECRET_PATH'), override=True)
