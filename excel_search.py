# excel_search.py

import os
import pandas as pd
import json
from typing import Optional, Union

EXCEL_FILES_PATH = r"./"  

def get_excel_path(filename: str) -> str:
    if os.path.isabs(filename):
        return filename

    if EXCEL_FILES_PATH is None:
        raise ValueError("Must provide absolute path if EXCEL_FILES_PATH is not set.")

    return os.path.join(EXCEL_FILES_PATH, filename)

def fast_search_in_excel(
    filepath: str,
    sheet_name: str,
    keyword: Union[str, int],
    usecols: Optional[str] = None,
    case_sensitive: bool = False
) -> str:
    try:
        full_path = get_excel_path(filepath)
        df = pd.read_excel(
            full_path,
            sheet_name=sheet_name,
            usecols=usecols,
            dtype=str,
            engine="openpyxl"
        )

        key_str = str(keyword)
        if not case_sensitive:
            key_str = key_str.lower()

        def row_matches(row):
            for cell in row:
                cell_str = str(cell)
                if not case_sensitive:
                    cell_str = cell_str.lower()
                if key_str in cell_str:
                    return True
            return False

        mask = df.apply(row_matches, axis=1)
        matched = df[mask]

        if matched.empty:
            return "No matching rows found."

        return matched.to_json(orient="records", force_ascii=False, indent=2)

    except Exception as e:
        return f"Error: {e}"
