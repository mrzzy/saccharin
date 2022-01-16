#
# sweetener
# blood sugar data cleaning & export
#

import numpy as np
import pandas as pd
from argparse import ArgumentParser

if __name__ == "__main__":
    # parse script command line args
    parser = ArgumentParser(description="Prepare blood sugar data for DDD evaluation")
    parser.add_argument("csv_path", help="Path to the blood sugar data CSV export.")
    args = parser.parse_args()

    # read sugar data csv export
    col_dtypes = {
        "Tags": str,
        "Blood Sugar Measurement (mmol/L)": np.float64,
        "Insulin Injection Units (Pen)": np.float64,
        "Basal Injection Units": np.float64,
        "Insulin (Meal)": np.float64,
        "Insulin (Correction)": np.float64,
        "Meal Carbohydrates (Grams, Factor 1)": np.float64,
        "Meal Descriptions": str,
        "Activity Duration (Minutes)": np.float64,
        "Activity Description": str,
        "Note": str,
        "HbA1c (Percent)": np.float64,
        "Food type": str,
    }
    date_cols = ["Date", "Time"]

    sugar_df = pd.read_csv(
        args.csv_path,
        usecols=list(col_dtypes.keys()) + date_cols,
        dtype=col_dtypes,
        parse_dates=date_cols,
    )
    if not isinstance(sugar_df, pd.DataFrame):
        raise ValueError("Expected sugar_df to be a DataFrame")
