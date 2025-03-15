import pandas as pd
import subprocess
import os
from datetime import datetime, timedelta

LECMD_PATH = "LECmd\\LECmd.exe"
RECMD_PATH = "RECmd\\RECmd.exe"
PECMD_PATH = "PECmd\\PECmd.exe"

LECMD_INPUT_FILE = None
REGISTRY_INPUT_FILE = None
REGISTRY_BATCH_FILE = "RECmd\\BatchExamples\\SoftwareASEPs.reb"
PECMD_INPUT_FILE = None

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

LECMD_CSV_PATH = OUTPUT_DIR
REGISTRY_CSV_PATH = OUTPUT_DIR
PECMD_CSV_PATH = OUTPUT_DIR
FINAL_CSV_PATH = OUTPUT_DIR

PECMD_OUT_FILE_NAME = "prefetch_output.csv"
RECMD_OUT_FILE_NAME = "registry_output.csv"
LECMD_OUT_FILE_NAME = "link_output.csv"
FINAL_OUT_FILE_NAME = "timeline_output_merged.xlsx"

# Function to run external command
def run_command(command):
    try:
        print(f"Running: {command}")
        subprocess.run(command, shell=True, check=True)
        print("Command executed successfully!")
    except subprocess.CalledProcessError as e:
        print(f"Error running command: {e}")

# Function to convert timestamps to UTC+08:00
def convert_to_utc8(timestamp):
    try:
        dt = datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S")  # Adjust if needed
        dt_utc8 = dt + timedelta(hours=8)  # Convert to UTC+08:00
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

def getFileInputDirectory(evtx_dir, pf_dir, reg_dir):
    
    print("\n--- Detected Directory Files ---")
    print("EVTX dir: " + check_path_exists(evtx_dir))
    print("PF dir: " + check_path_exists(pf_dir))
    print("REG dir: " + check_path_exists(reg_dir))
    print("---       END      ---\n")

    if check_path_exists(pf_dir) != "Not Found":
        return 1
    else:
        return 0
    
def getFileInputNames():
    directory = os.path.dirname(os.path.abspath(__file__))

    amache_file = None
    ln_file = None
    reg_file = None

    for file in os.listdir(directory):
        if os.path.isfile(os.path.join(directory, file)):  # Ensure it's a file, not a folder
            if file.endswith(".lnk"):
                ln_file = file
            elif file.endswith(".hve"):
                amache_file = file
            elif "." not in file:
                reg_file = file

    print("\n--- Detected Files ---")
    print("Amache File:", amache_file if amache_file else "Not Found")
    print("HIVE File:", reg_file if reg_file else "Not Found")
    print("LINK File:", ln_file if ln_file else "Not Found")
    print("---       END      ---\n")

    return amache_file, reg_file, ln_file

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
    amache_command = f'"AmcacheParser.exe -f {PECMD_INPUT_FILE} --csv {PECMD_CSV_PATH} --csvf {PECMD_OUT_FILE_NAME}"'

    try:
        print(f"Running RECmd on {PECMD_INPUT_FILE}...")
        subprocess.run(amache_command, shell=True, check=True)
        print(f"Registry parsed successfully! Output saved to: {PECMD_CSV_PATH}")
    except subprocess.CalledProcessError as e:
        print(f"Error running PECmd: {e}")

# Function to run LECmd and process shortcut files
def run_lecmd(LNK_INPUT_DIR):
    lnk_command = f"LECmd.exe -d {LNK_INPUT_DIR} --csv {LECMD_CSV_PATH} --csvf {LECMD_OUT_FILE_NAME}"

    try:
        print(f"Running RECmd on {LNK_INPUT_DIR}...")
        subprocess.run(lnk_command, shell=True, check=True)
        print(f"Registry parsed successfully! Output saved to: {LECMD_CSV_PATH}")
    except subprocess.CalledProcessError as e:
        print(f"Error running LECmd: {e}")

# Main function to generate a forensic timeline
def create_timeline(df_recmd, df_amcache, df_lecmd, output_csv):
    final_df = pd.concat([df_recmd, df_amcache, df_lecmd], ignore_index=True)
    final_df = final_df.sort_values(by='Timestamp', ascending=True)

    final_df.to_csv(output_csv, index=False)
    print(f"Timeline saved to {output_csv}")

# Example usage: Replace with actual paths
def main():
    run_recmd()
    # create_timeline("Registry_Folder", "Amcache.hve", "LNK_Folder", "Final_Timeline_UTC+0800.csv")

if __name__ == "__main__":
    main()