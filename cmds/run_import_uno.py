"""
Recursively add 'import uno' to all __init__.py files in the given directory if it does not already exist.
"""

import os
from pathlib import Path


def add_import_uno_to_init_files(root_dir):
    """
    Recursively add 'import uno' to all __init__.py files in the given directory if it does not already exist.

    Args:
        root_dir (str): The root directory to start the search.
    """
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename == "__init__.py":
                file_path = os.path.join(dirpath, filename)
                add_import_uno(file_path)


def add_import_uno(file_path):
    """
    Add 'import uno' to the given file if it does not already exist.

    Args:
        file_path (str): The path to the file to modify.
    """
    with open(file_path, "r") as file:
        lines = file.readlines()

    # Check if 'import uno' already exists
    if any("import uno" in line for line in lines):
        return

    # append 'import uno' to the file
    lines.append("import uno\n")

    with open(file_path, "w") as file:
        file.writelines(lines)


def find_ooodev_directory(start_path):
    """
    Walk up the directory tree until 'ooodev' directory is found.

    Args:
        start_path (str): The starting path to begin the search.

    Returns:
        str: The path to the 'ooodev' directory if found, otherwise None.
    """
    current_path = Path(start_path).resolve()

    while current_path != current_path.parent:
        if (current_path / "ooodev").is_dir():
            return str(current_path / "ooodev")
        current_path = current_path.parent

    return None


if __name__ == "__main__":
    root_directory = find_ooodev_directory(os.getcwd())
    # print(root_directory)
    add_import_uno_to_init_files(root_directory)
