# main.py

import json
import os
import logging
import argparse
import sys
from modules.file_loader import load_file
from modules.column_mapper import map_columns
from modules.formatter import apply_formatting
from modules.output_handler import save_output
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import pandas as pd
from modules.overview.overview_handler import create_overview

logging.basicConfig(
    filename='formatter.log',
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

class ConfigHandler(FileSystemEventHandler):
    def __init__(self, config_path):
        super().__init__()
        self.config_path = config_path
        self.config = self.load_config()

    def load_config(self):
        with open(self.config_path, "r") as f:
            config = json.load(f)
        logging.info(f"Configuration loaded from {self.config_path}")
        return config

    def on_modified(self, event):
        if event.src_path == os.path.abspath(self.config_path):
            try:
                self.config = self.load_config()
                logging.info(f"Configuration reloaded due to modification in {self.config_path}")
            except Exception as e:
                logging.exception("Failed to reload configuration.")

def parse_arguments():
    parser = argparse.ArgumentParser(description="Excel/CSV File Formatter")
    parser.add_argument(
        'file_path',
        type=str,
        help='Path to the input CSV, Excel, or JSON file.'
    )
    parser.add_argument(
        '-c', '--config',
        type=str,
        default='config/column_rules.json',
        help='Path to the column rules configuration JSON file.'
    )
    parser.add_argument(
        '-o', '--output',
        type=str,
        help='Path to save the formatted Excel file.'
    )
    return parser.parse_args()

def start_config_watcher(config_path, config_handler):
    observer = Observer()
    observer.schedule(config_handler, path=os.path.dirname(os.path.abspath(config_path)), recursive=False)
    observer.start()
    logging.info(f"Started configuration watcher on {config_path}")
    return observer

def main_process(file_path, innography_columns, formatting_rules, output_path):
    try:
        logging.info(f"Starting processing for file: {file_path}")
        
        data, workbook = load_file(file_path)

        data = map_columns(data, innography_columns)

        if "Type (Grant/Application)" in data.columns:
            data["Type (Grant/Application)"] = data["Type (Grant/Application)"].astype(str).apply(
                lambda x: "Application" if "application" in x.lower() else ("Grant" if "grant" in x.lower() else x)
            )
            logging.info('Standardized "Type (Grant/Application)" column values.')

        if "Publication Country" in data.columns:
            pub_country_idx = data.columns.get_loc("Publication Country")
            data.insert(pub_country_idx, "Jurisdiction (US, Non-US)", "")

            data["Jurisdiction (US, Non-US)"] = data["Publication Country"].apply(
                lambda x: "US" if str(x).strip().upper() == "US" else "Non-US"
            )
            logging.info('Inserted "Jurisdiction (US, Non-US)" column based on "Publication Country".')

        formatted_workbook = apply_formatting(data, workbook, innography_columns, formatting_rules)

        save_output(formatted_workbook, output_path)

        create_overview(output_path) 

        logging.info(f"Processing completed for file: {file_path}")

    except Exception as e:
        logging.error(f"Processing failed for file: {file_path}. Error: {e}")
        print(f"An error occurred during processing: {e}")

def main():
    args = parse_arguments()
    config_path = args.config
    file_path = args.file_path
    output_path = args.output
    
    if not os.path.isfile(file_path):
        logging.error(f"Input file does not exist: {file_path}")
        print(f"Input file does not exist: {file_path}")
        sys.exit(1)

    if not os.path.isfile(config_path):
        logging.error(f"Configuration file does not exist: {config_path}")
        print(f"Configuration file does not exist: {config_path}")
        sys.exit(1)
    try:
        with open(config_path, "r") as f:
            config = json.load(f)
        logging.info(f"Loaded configuration from {config_path}")
    except Exception as e:
        logging.exception("Failed to load configuration.")
        print(f"Failed to load configuration: {e}")
        sys.exit(1)

    config_handler = ConfigHandler(config_path)

    observer = start_config_watcher(config_path, config_handler)

    try:
        if not output_path:
            base, ext = os.path.splitext(file_path)
            output_path = f"{base}_DONE.xlsx"

        main_process(
            file_path,
            config_handler.config["InnographyColumns"],
            config_handler.config["FormattingRules"],
            output_path
        )

    finally:
        observer.stop()
        observer.join()

if __name__ == "__main__":
    main()
