# make sure pytest see root dir.
import os
import sys
from pathlib import Path
def insert_root_path():
    root = Path(__file__).parent
    sys.path.insert(0, str(root))

if os.name == 'nt':
    insert_root_path()