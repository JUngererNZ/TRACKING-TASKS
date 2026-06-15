#!/usr/bin/env python3
"""
Directory Structure to JSON

This script walks through a given directory and exports its
hierarchy (folders and files) to a JSON file.
"""

import json
import argparse
from pathlib import Path
from datetime import datetime

def get_directory_structure(path: Path) -> dict:
    """
    Recursively build a dictionary representing the directory tree.

    Args:
        path (Path): Root directory to scan.

    Returns:
        dict: Dictionary with keys: name, path, type, modified, and children (for directories).
    """
    if not path.exists():
        raise FileNotFoundError(f"Path does not exist: {path}")
    if not path.is_dir():
        raise NotADirectoryError(f"Path is not a directory: {path}")

    try:
        stat = path.stat()
        modified = datetime.fromtimestamp(stat.st_mtime).isoformat()
    except PermissionError:
        modified = None

    node = {
        "name": path.name,
        "path": str(path.absolute()),
        "type": "directory",
        "modified": modified,
        "children": []
    }

    try:
        for item in sorted(path.iterdir()):
            if item.is_dir():
                # Recursively process subdirectories
                node["children"].append(get_directory_structure(item))
            else:
                # Add file information
                try:
                    file_stat = item.stat()
                    file_modified = datetime.fromtimestamp(file_stat.st_mtime).isoformat()
                except PermissionError:
                    file_modified = None

                node["children"].append({
                    "name": item.name,
                    "path": str(item.absolute()),
                    "type": "file",
                    "size": file_stat.st_size if 'file_stat' in locals() else None,
                    "modified": file_modified
                })
    except PermissionError:
        # If we can't read the directory, leave children empty
        pass

    return node

def main():
    parser = argparse.ArgumentParser(
        description="Export directory structure (parent + subfolders) to a JSON file."
    )
    parser.add_argument(
        "--path", "-p",
        type=str,
        default=".",
        help="Root directory to scan (default: current working directory)."
    )
    parser.add_argument(
        "--output", "-o",
        type=str,
        default="directory_structure.json",
        help="Output JSON file name (default: directory_structure.json)."
    )
    args = parser.parse_args()

    root_path = Path(args.path).resolve()

    try:
        print(f"Scanning directory: {root_path}")
        structure = get_directory_structure(root_path)
    except (FileNotFoundError, NotADirectoryError, PermissionError) as e:
        print(f"Error: {e}")
        return 1

    try:
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(structure, f, indent=2, ensure_ascii=False)
        print(f"Directory structure saved to: {args.output}")
    except IOError as e:
        print(f"Error writing JSON file: {e}")
        return 1

    return 0

if __name__ == "__main__":
    exit(main())