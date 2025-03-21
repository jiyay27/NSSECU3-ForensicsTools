import os
import time
import pandas as pd
import subprocess
from datetime import datetime, timedelta, timezone
utc_now = datetime.now(timezone.utc)
day_scope = utc_now - timedelta(days=40)
day_scope = day_scope.strftime("%Y-%m-%d %H:%M:%S")

EVTXECMD_PATH = "EvtxECmd\\EvtxECmd.exe"
RECMD_PATH = "RECmd\\RECmd.exe"
PECMD_PATH = "PECmd\\PECmd.exe"

EVTX_INPUT_FILE = None
REGISTRY_INPUT_FILE = None
REGISTRY_BATCH_FILE = "RECmd\\BatchExamples\\SoftwareASEPs.reb"
PREFETCH_INPUT_FILE = None

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

EVTX_CSV_PATH = OUTPUT_DIR
REGISTRY_CSV_PATH = OUTPUT_DIR
PREFETCH_CSV_PATH = OUTPUT_DIR
FINAL_CSV_PATH = OUTPUT_DIR

EVTX_OUT_FILE_NAME = "evtx_output.csv"
RECMD_OUT_FILE_NAME = "registry_output.csv"
PECMD_OUT_FILE_NAME = "prefetch_output.csv"
PECMD2_OUT_FILE_NAME = "prefetch_output_Timeline.csv"
FINAL_OUT_FILE_NAME = "forensic_output_merged.xlsx"

# Make sure the output directory exists
os.makedirs(OUTPUT_DIR, exist_ok=True)

def run_evtxecmd(EVTX_INPUT_FILE):
    command = f'"{EVTXECMD_PATH}" -f "{EVTX_INPUT_FILE}" --sd "{day_scope}" --csv {EVTX_CSV_PATH} --csvf {EVTX_OUT_FILE_NAME}' #--maps "{EVTXECMD_PATH}\\Maps"'
    try:
        print(f'"Running EvtxECmd to parse {EVTX_INPUT_FILE}..."')
        subprocess.run(command, shell=True, check=True)

        start_time = time.time()
        while not os.path.exists(EVTX_CSV_PATH):
            if time.time() - start_time > 10:
                raise TimeoutError(f"File {EVTX_CSV_PATH} was not created in time.")
            time.sleep(1)

        print(f'"\n{EVTX_INPUT_FILE} successfully parsed! CSV saved at: {EVTX_CSV_PATH}"')

    except subprocess.CalledProcessError as e:
        print(f"Error running EvtxECmd: {e}")
    except TimeoutError as e:
        print(f"Timeout error: {e}")

def run_recmd(REGISTRY_INPUT_FILE, REGISTRY_BATCH_FILE):
    recmd_command = f'"{RECMD_PATH}" -f "{REGISTRY_INPUT_FILE}" --bn "{REGISTRY_BATCH_FILE}" --csv {REGISTRY_CSV_PATH} --csvf {RECMD_OUT_FILE_NAME}'

    try:
        print(f"Running RECmd on {REGISTRY_INPUT_FILE}...")
        subprocess.run(recmd_command, shell=True, check=True)
        print(f"Registry parsed successfully! Output saved to: {REGISTRY_CSV_PATH}")
    except subprocess.CalledProcessError as e:
        print(f"Error running RECmd: {e}")

def run_pecmd(PREFETCH_INPUT_FILE):
    command = f'"{PECMD_PATH}" -f "{PREFETCH_INPUT_FILE}" --csv {PREFETCH_CSV_PATH} --csvf {PECMD_OUT_FILE_NAME}'

    try:
        print(f"Running PECmd on {PREFETCH_INPUT_FILE}...")
        subprocess.run(command, shell=True, check=True)

        start_time = time.time()
        while not os.path.exists(PREFETCH_CSV_PATH):
            if time.time() - start_time > 10:
                raise TimeoutError(f"File {PREFETCH_CSV_PATH} was not created in time.")
            time.sleep(1)

        print(f"Prefetch file {PREFETCH_INPUT_FILE} successfully parsed! CSV saved at: {PREFETCH_CSV_PATH}")

    except subprocess.CalledProcessError as e:
        print(f"Error running PECmd: {e}")
    except TimeoutError as e:
        print(f"Timeout error: {e}")

def create_forensic_timeline():
    print("\nMerging forensic data into a single timeline...")
    
    # Read each csv file into a dataframe
    evtx_df = pd.read_csv(EVTX_OUT_FILE_NAME, low_memory=False) if os.path.exists(EVTX_CSV_PATH) else pd.DataFrame()
    registry_df = pd.read_csv(RECMD_OUT_FILE_NAME, low_memory=False) if os.path.exists(REGISTRY_CSV_PATH) else pd.DataFrame()
    prefetch_df = pd.read_csv(PECMD_OUT_FILE_NAME, low_memory=False) if os.path.exists(PREFETCH_CSV_PATH) else pd.DataFrame()
    prefetch2_df = pd.read_csv(PECMD2_OUT_FILE_NAME, low_memory=False) if os.path.exists(PREFETCH_CSV_PATH) else pd.DataFrame()

    # If final output file already exists, remove it
    if os.path.exists(FINAL_OUT_FILE_NAME):
        print(f"\nFile '{FINAL_OUT_FILE_NAME}' already exists. Overwriting...")
        os.remove(FINAL_OUT_FILE_NAME)

    time.sleep(1)

    # Write separate xlsx files to final compiled xlsx file
    with pd.ExcelWriter(FINAL_OUT_FILE_NAME, engine="openpyxl") as writer:
        if not evtx_df.empty:
            evtx_df.to_excel(writer, sheet_name="EVTX_Logs", index=False)
        if not registry_df.empty:
            registry_df.to_excel(writer, sheet_name="Registry_Logs", index=False)
        if not prefetch_df.empty:
            prefetch_df.to_excel(writer, sheet_name="Prefetch_Logs", index=False)
        if not prefetch2_df.empty:
            prefetch2_df.to_excel(writer, sheet_name="Prefetch_Logs2", index=False)

    time.sleep(1)

    # If final output file was created successfully, print success message
    if os.path.exists(FINAL_OUT_FILE_NAME):
        print(f"Forensic timeline created successfully! Saved at: {FINAL_CSV_PATH}")
    else:
        print("Error creating forensic timeline")

    # If the file exists, remove it since we already have the Final output file
    if os.path.exists(EVTX_OUT_FILE_NAME):
        os.remove(EVTX_OUT_FILE_NAME)
    if os.path.exists(RECMD_OUT_FILE_NAME):
        os.remove(RECMD_OUT_FILE_NAME)
    if os.path.exists(PECMD_OUT_FILE_NAME):
        os.remove(PECMD_OUT_FILE_NAME)
    if os.path.exists(PECMD2_OUT_FILE_NAME):
        os.remove(PECMD2_OUT_FILE_NAME)

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

    evtx_file = None
    pf_file = None
    reg_file = None

    for file in os.listdir(directory):
        if os.path.isfile(os.path.join(directory, file)):  # Ensure it's a file, not a folder
            if file.endswith(".evtx"):
                evtx_file = file
            elif file.endswith(".pf"):
                pf_file = file
            elif "." not in file:
                reg_file = file

    print("\n--- Detected Files ---")
    print("EVTX File:", evtx_file if evtx_file else "Not Found")
    print("PF File:", pf_file if pf_file else "Not Found")
    print("REG File:", reg_file if reg_file else "Not Found")
    print("---       END      ---\n")

    return evtx_file, pf_file, reg_file

def getUserInputBatchFilters():
    getBatchFileFilters()
    bn_filter = input("Batch File Filter: ")
    if bn_filter == "0":
        print("Exiting...")
        os._exit(0);
    return (".\\RECmd\\BatchExamples\\" + bn_filter)

def runAllTools1(EVTX_INPUT_FILE, PREFETCH_INPUT_FILE,  REGISTRY_INPUT_FILE, REGISTRY_BATCH_FILE):
    run_evtxecmd(EVTX_INPUT_FILE)
    run_recmd(REGISTRY_INPUT_FILE, REGISTRY_BATCH_FILE)
    run_pecmd(PREFETCH_INPUT_FILE)
    create_forensic_timeline()

def header():
    print("--- NSSECU3 Forensics Tool ---")
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
        EVTX_INPUT_FILE, PREFETCH_INPUT_FILE,  REGISTRY_INPUT_FILE = getFileInputNames()
        time.sleep(1)
        REGISTRY_BATCH_FILE = getUserInputBatchFilters()
        runAllTools1(EVTX_INPUT_FILE, PREFETCH_INPUT_FILE,  REGISTRY_INPUT_FILE, REGISTRY_BATCH_FILE)
        
    if val == "2":
        print("\nNOTE: Ensure that LIVE files are not scanned as the EvtxECmd Tool does not allow it.")
        print("Also, write the directory name with no quotation marks (\"\") or double backslashes (\\\\).\n")

        print("--- Specify User Directory ---")

        # "C:\Windows\System32\winevt\Logs\Internet Explorer.evtx" - sample dir
        evtx_dir = input("EVTX Directory: ")

        # "C:\Windows\Prefetch\WSL.EXE-1B26B53B.pf" - sample dir
        pf_dir = input("PF Directory: ")

        # "C:\Windows\System32\config\SOFTWARE" - sample dir
        reg_dir = input("REG Directory: ")

        print("Scanning directory for specific files...")
        time.sleep(1)
        out = getFileInputDirectory(evtx_dir, pf_dir, reg_dir)
        time.sleep(1)

        if out == 1:
            REGISTRY_BATCH_FILE = getUserInputBatchFilters()
        else:
            REGISTRY_BATCH_FILE = 0
            print("No registry batch file found.")
            print("Filters will not be used.")

        runAllTools1(evtx_dir, pf_dir,  reg_dir, REGISTRY_BATCH_FILE)

    if val == "3":
        print("Exiting...")
        os._exit(0);

if __name__ == "__main__":
    main()