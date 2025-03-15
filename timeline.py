import time
import pandas as pd
import subprocess
import os
from datetime import datetime, timedelta

LECMD_PATH = "LECmd\\LECmd.exe"
RECMD_PATH = "RECmd\\RECmd.exe"
PECMD_PATH = "PECmd\\PECmd.exe"

LECMD_INPUT_FILE = None # folder for multiple files
REGISTRY_INPUT_FILE = None # single file
REGISTRY_BATCH_FILE = "RECmd\\BatchExamples\\SoftwareASEPs.reb"
PECMD_INPUT_FILE = None # single file

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

LECMD_CSV_PATH = OUTPUT_DIR
REGISTRY_CSV_PATH = OUTPUT_DIR
PECMD_CSV_PATH = OUTPUT_DIR
FINAL_CSV_PATH = OUTPUT_DIR

LECMD_OUT_FILE_NAME = "link_output.csv"
RECMD_OUT_FILE_NAME = "registry_output.csv"
PECMD_OUT_FILE_NAME = "prefetch_output.csv"
FINAL_OUT_FILE_NAME = "timeline_output_merged.xlsx"

# Function to run external command
def run_command(command):
    try:
        print(f"Running: {command}")
        subprocess.run(command, shell=True, check=True)
        print("Command executed successfully!")
    except subprocess.CalledProcessError as e:
        print(f"Error running command: {e}")

def convert_to_utc(timestamp):
        try:
            dt = datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S")
            return dt.strftime("%Y-%m-%d %H:%M:%S")  # Normalize to UTC+0000
        except:
            return timestamp  # Return as-is if parsing fails

def convert_to_utc8(timestamp):
    try:
        dt = datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S")
        dt_utc8 = dt + timedelta(hours=8)
        return dt_utc8.strftime("%Y-%m-%d %H:%M:%S")
    except:
        return timestamp  # Return as-is if parsing fails
    
def getBatchFileFilters():
    directory = ".\\RECmd\\BatchExamples\\"
    batch_files = []

    for file in os.listdir(directory):
        if file.endswith(".reb"):
            batch_files.append(file)

    print("--- Batch File Filter Options ---")
    for file in batch_files:
        print(file);
    print("---            END            ---\n")

def getUserInputBatchFilters():
    getBatchFileFilters()
    bn_filter = input("Batch File Filter: ")
    if bn_filter == "0":
        print("Exiting...")
        os._exit(0);
    return (".\\RECmd\\BatchExamples\\" + bn_filter)

def check_path_exists(path):
    if os.path.exists(path):
        if os.path.isfile(path):
            return path
        elif os.path.isdir(path):
            return path
    return "Not Found"

def getFileInputDirectory(lnk_dir, reg_dir, pf_dir):
    
    print("\n--- Detected Directory Files ---")
    print("LNK dir: " + check_path_exists(lnk_dir))
    print("REG dir: " + check_path_exists(reg_dir))
    print("PF dir: " + check_path_exists(pf_dir))
    print("---       END      ---\n")

    if check_path_exists(pf_dir) != "Not Found":
        return 1
    else:
        return 0
    
def getFileInputNames():
    directory = os.path.dirname(os.path.abspath(__file__))

    pref_file = None
    ln_file = None
    reg_file = None

    for file in os.listdir(directory):
        if os.path.isfile(os.path.join(directory, file)):  # Ensure it's a file, not a folder
            if file.endswith(".lnk"):
                ln_file = OUTPUT_DIR
            elif file.endswith(".pf"):
                pref_file = file
            elif "." not in file:
                reg_file = file

    print("\n--- Detected Files ---")
    print("LINK File:", OUTPUT_DIR if ln_file else "Not Found")
    print("HIVE File:", reg_file if reg_file else "Not Found")
    print("PREF File:", pref_file if pref_file else "Not Found")
    print("---       END      ---\n")

    return ln_file, reg_file, pref_file

# Function to run RECmd and process registry data
def run_recmd(REGISTRY_INPUT_FILE, REGISTRY_BATCH_FILE):
    recmd_command = f'"{RECMD_PATH}" -f "{REGISTRY_INPUT_FILE}" --bn "{REGISTRY_BATCH_FILE}" --csv {REGISTRY_CSV_PATH} --csvf {RECMD_OUT_FILE_NAME}'

    try:
        print(f"Running RECmd on {REGISTRY_INPUT_FILE}...")
        subprocess.run(recmd_command, shell=True, check=True)
        print(f"Registry parsed successfully! Output saved to: {REGISTRY_CSV_PATH}")
    except subprocess.CalledProcessError as e:
        print(f"Error running RECmd: {e}")

# Function to run AmcacheParser and process execution history
def run_pecmd(PECMD_INPUT_FILE):
    pecmd_command = f'"{PECMD_PATH} -f {PECMD_INPUT_FILE} --csv {PECMD_CSV_PATH} --csvf {PECMD_OUT_FILE_NAME}"'

    try:
        print(f"Running PECmd on {PECMD_INPUT_FILE}...")
        subprocess.run(pecmd_command, shell=True, check=True)

        start_time = time.time()
        while not os.path.exists(PECMD_CSV_PATH):
            if time.time() - start_time > 10:
                raise TimeoutError(f"File {PECMD_CSV_PATH} was not created in time.")
            time.sleep(1)

        print(f"Prefetch file {PECMD_INPUT_FILE} successfully parsed! CSV saved at: {PECMD_CSV_PATH}")

    except subprocess.CalledProcessError as e:
        print(f"Error running PECmd: {e}")
    except TimeoutError as e:
        print(f"Timeout error: {e}")

# Function to run LECmd and process shortcut files
def run_lecmd(LNK_INPUT_DIR):
    lnk_command = f'"{LECMD_PATH}" -d "{LNK_INPUT_DIR}" --csv "{LECMD_CSV_PATH}" --csvf "{LECMD_OUT_FILE_NAME}"'

    try:
        print(f"Running LECmd on {LNK_INPUT_DIR}...")
        subprocess.run(lnk_command, shell=True, check=True)
        print(f"Registry parsed successfully! Output saved to: {LECMD_CSV_PATH}")
    except subprocess.CalledProcessError as e:
        print(f"Error running LECmd: {e}")

def adjust_column_widths(writer, df, sheet_name):
    worksheet = writer.sheets[sheet_name]
    for i, col in enumerate(df.columns):
        max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, max_len)

def create_timeline(recmd_csv, pecmd_csv, lecmd_csv, output_csv):
    # Read CSV files if they exist
    df_recmd = pd.read_csv(recmd_csv) if recmd_csv and os.path.exists(recmd_csv) else pd.DataFrame()
    df_pecmd = pd.read_csv(pecmd_csv) if pecmd_csv and os.path.exists(pecmd_csv) else pd.DataFrame()
    df_lecmd = pd.read_csv(lecmd_csv) if lecmd_csv and os.path.exists(lecmd_csv) else pd.DataFrame()
    
    # Merge DataFrames
    final_df = pd.concat([df_recmd, df_pecmd, df_lecmd], ignore_index=True)
    
    # Check if 'Timestamp' column exists
    if 'Timestamp' not in final_df.columns:
        print("Error: 'Timestamp' column not found in the CSV files.")
        return
    
    # Convert timestamps
    final_df["UTC+0000"] = final_df["Timestamp"].apply(convert_to_utc)
    final_df["UTC+0800"] = final_df["UTC+0000"].apply(convert_to_utc8)
    
    # Sort by UTC+0000
    final_df = final_df.sort_values(by="UTC+0000", ascending=True)
    
    # Save the final timeline in Excel with adjusted column sizes
    with pd.ExcelWriter(output_csv, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, sheet_name='Timeline', index=False)
        adjust_column_widths(writer, final_df, 'Timeline')
    
    print(f"Timeline saved to {output_csv}")
    
    # Delete the original CSV files
    for csv_file in [recmd_csv, pecmd_csv, lecmd_csv]:
        if csv_file and os.path.exists(csv_file):
            os.remove(csv_file)
            print(f"Deleted {csv_file}")

def runAllTools1(LECMD_INPUT_FILE, PECMD_INPUT_FILE, REGISTRY_INPUT_FILE, REGISTRY_BATCH_FILE):
    run_lecmd(LECMD_INPUT_FILE)
    run_recmd(REGISTRY_INPUT_FILE, REGISTRY_BATCH_FILE)
    run_pecmd(PECMD_INPUT_FILE)
    create_timeline(REGISTRY_CSV_PATH, PECMD_CSV_PATH, LECMD_CSV_PATH, FINAL_CSV_PATH)


def header():
    print("--- NSSECU3 Timeline Tool ---")
    print("[1] Scan Files in Current Directory")
    print("[2] Scan Files in User Specified Directory")
    print("[3] Exit")
    print("--- ---------------------- ---")



def main():
    header()
    val = input("Input: ")

    if val == "1":
        print("\nScanning files in current directory...")
        time.sleep(1)
        LECMD_INPUT_FILE, REGISTRY_INPUT_FILE,  PECMD_INPUT_FILE = getFileInputNames()
        print(LECMD_INPUT_FILE)
        time.sleep(1)
        REGISTRY_BATCH_FILE = getUserInputBatchFilters()
        runAllTools1(LECMD_INPUT_FILE, PECMD_INPUT_FILE, REGISTRY_INPUT_FILE, REGISTRY_BATCH_FILE)
    
    if val == "2":
        print("\nNOTE: Ensure that LIVE files are not scanned as the EvtxECmd Tool does not allow it.")
        print("Also, write the directory name with no quotation marks (\"\") or double backslashes (\\\\).\n")

    # create_timeline("Registry_Folder", "Amcache.hve", "LNK_Folder", "Final_Timeline_UTC+0800.csv")

if __name__ == "__main__":
    main()