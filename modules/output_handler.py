# modules/output_handler.py

import logging
from openpyxl import Workbook

def save_output(workbook: Workbook, output_path: str):
    try:
        workbook.save(output_path)
        logging.info(f"Formatted file saved successfully to {output_path}")
        print(f"Formatted file saved successfully to {output_path}")
    except Exception as e:
        logging.exception("Error saving the output file.")
        raise e
