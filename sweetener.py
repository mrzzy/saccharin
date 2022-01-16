#
# sweetener
# blood sugar data cleaning & export
#

import numpy as np
import pandas as pd
from datetime import datetime
from argparse import ArgumentParser

if __name__ == "__main__":
    # parse script command line args
    parser = ArgumentParser(description="Prepare blood sugar data for DDD evaluation")
    parser.add_argument("csv_path", help="Path to the blood sugar data CSV export.")
    parser.add_argument("--start-from", help="Starting date in format DD/MM/YYYY. "
                        "Data from the starting date onwards is included.",
                        type=(lambda date_str: datetime.strptime(date_str, "%d/%m/%Y")))
    args = parser.parse_args()

    # read sugar data csv export
    col_dtypes = {
        "Date": str,
        "Time": str,
        "Tags": str,
        "Blood Sugar Measurement (mmol/L)": np.float64,
        "Basal Injection Units": np.float64,
        "Insulin (Meal)": np.float64,
        "Insulin (Correction)": np.float64,
        "Meal Carbohydrates (Grams, Factor 1)": np.float64,
        "Meal Descriptions": str,
        "Note": str,
    }

    sugar_df = pd.read_csv(
        args.csv_path,
        usecols=list(col_dtypes.keys()),
        dtype=col_dtypes,
    )
    if not isinstance(sugar_df, pd.DataFrame):
        raise ValueError("Expected sugar_df to be a DataFrame")

    # parse date & time columns
    sugar_df["Timestamp"] = pd.to_datetime(
        sugar_df["Date"] + " " + sugar_df["Time"]
    )
    sugar_df["Date"] = pd.to_datetime(sugar_df["Date"])
    sugar_df["Time"] = pd.to_datetime(sugar_df["Time"]).apply((lambda dt: dt.time())) # type: ignore

    # filter data to only include entries from start date onwards
    if args.start_from is not None:
        sugar_df = sugar_df[sugar_df["Date"] >= args.start_from]

    # add hypo / hyperglycemia features
    sugar_df["Hyperglycemia"] =  sugar_df["Blood Sugar Measurement (mmol/L)"] > 10.0
    sugar_df["Hypoglycemia"] = sugar_df["Blood Sugar Measurement (mmol/L)"] < 4.0

    print(sugar_df[sugar_df["Hypoglycemia"]])
