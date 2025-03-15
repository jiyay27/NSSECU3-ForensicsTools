import pandas as pd
from datetime import datetime
import pytz

def normalize_timestamps(df, timestamp_cols, tz_from, tz_to):
    tz_from = pytz.timezone(tz_from)
    tz_to = pytz.timezone(tz_to)

    for col in timestamp_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')  # Convert to datetime
            df[col] = df[col].dt.tz_localize(tz_from, ambiguous='NaT').dt.tz_convert(tz_to).dt.strftime("%Y-%m-%d %H:%M:%S %z")
    
    return df

def merge_csvs(csv1, csv2, csv3, output_file):
    # Load CSVs
    df1 = pd.read_csv(csv1, usecols=["RunTime", "ExecutableName"])
    df2 = pd.read_csv(csv2, usecols=["BagPath", "AbsolutePath", "CreatedOn", "ModifiedOn", "AccessedOn", "LastWriteTime", "FirstInteracted", "LastInteracted", "HasExplored", "Miscellaneous"])
    df3 = pd.read_csv(csv3, usecols=["SourceFile", "SourceCreated", "SourceModified", "SourceAccessesd", "TargetCreated", "TargetModified", "TargetAcessed", "LocalPath", "MachineMACAddress", "TrackerCreatedOn"])


    # Define a unified column structure
    unified_columns = [
        "AccessTime", "ExecutableName", 
    ]

    # Normalize column names in each DataFrame
    # df1 = df1.rename(columns={"RunTime": "RunTime", "ExecutableName": "ExecutableName"})
    df2 = df2.rename(columns={"CreatedOn": "CreatedOn", "ModifiedOn": "ModifiedOn", "AccessedOn": "AccessedOn", "LastWriteTime": "LastWriteTime",
                               "FirstInteracted": "FirstInteracted", "LastInteracted": "LastInteracted", "HasExplored": "HasExplored"})
    df3 = df3.rename(columns={"SourceCreated": "SourceCreated", "SourceModified": "SourceModified", "SourceAccessesd": "SourceAccessed",
                               "TargetCreated": "TargetCreated", "TargetModified": "TargetModified", "TargetAcessed": "TargetAccessed",
                               "LocalPath": "LocalPath", "MachineMACAddress": "MachineMACAddress"})

    # Fill missing columns with NaN
    for df in [df1, df2, df3]:
        for col in unified_columns:
            if col not in df.columns:
                df[col] = None

    # Merge all data
    merged_df = pd.concat([df1, df2, df3], ignore_index=True)

    # Define all timestamp columns
    timestamp_cols = [
        "RunTime", "CreatedOn", "ModifiedOn", "AccessedOn", "LastWriteTime",
        "FirstInteracted", "LastInteracted", "SourceCreated", "SourceModified",
        "SourceAccessed", "TargetCreated", "TargetModified", "TargetAccessed"
    ]

    # Normalize timestamps to UTC+0000
    merged_df_utc = normalize_timestamps(merged_df.copy(), timestamp_cols, "UTC", "UTC")

    # Normalize timestamps to UTC+0800
    merged_df_utc8 = normalize_timestamps(merged_df.copy(), timestamp_cols, "UTC", "Asia/Singapore")

    # Save to an Excel file with two sheets (UTC+0000 and UTC+0800)
    with pd.ExcelWriter(output_file) as writer:
        merged_df_utc.to_excel(writer, sheet_name="UTC+0000", index=False)
        merged_df_utc8.to_excel(writer, sheet_name="UTC+0800", index=False)

    print(f"Final timeline saved to {output_file} with two timezone sheets.")

# Example usage
merge_csvs("link_output.csv", "John_UsrClass_shellb_output.csv", "prefetch_output_Timeline.csv", "final_output.xlsx")
