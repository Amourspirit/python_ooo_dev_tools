#!/usr/bin/env python

import os
from pathlib import Path


def count_py_lines(root_dir):
    total_lines = 0
    for dirpath, dirnames, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.endswith(".py"):
                file_path = os.path.join(dirpath, filename)
                with open(file_path, "r") as file:
                    lines = file.readlines()
                    total_lines += len(lines)
    print(f"Total number of Python code lines: {total_lines}")


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
    project_directory = find_ooodev_directory(str(Path.cwd()))
    count_py_lines(project_directory)
