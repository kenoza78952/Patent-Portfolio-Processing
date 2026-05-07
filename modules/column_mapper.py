# modules/column_mapper.py

import logging
import pandas as pd

def map_columns(data: pd.DataFrame, innography_columns: dict) -> pd.DataFrame:
    try:
        column_mapping = {}
        for final_name, details in innography_columns.items():
            for alias in details.get("aliases", []):
                column_mapping[alias] = final_name

        logging.info(f"Column mapping: {column_mapping}")

        data = data.rename(columns=column_mapping)

        logging.debug(f"After renaming, column names in DataFrame: {data.columns.tolist()}")

        ordered_columns = sorted(
            [(name, details["order"]) for name, details in innography_columns.items()],
            key=lambda x: x[1]
        )
        ordered_column_names = [col[0] for col in ordered_columns]

        final_columns = [col for col in ordered_column_names if col in data.columns]
        extra_columns = [col for col in data.columns if col not in final_columns]

        ordered_data = data[final_columns + extra_columns]

        logging.debug(f"Final ordered column names in DataFrame: {ordered_data.columns.tolist()}")

        return ordered_data

    except Exception as e:
        logging.exception("Error in map_columns function.")
        raise e
