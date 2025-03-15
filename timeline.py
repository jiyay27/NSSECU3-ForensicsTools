import time
import subprocess
import os

LECMD_PATH = "LECmd\\LECmd.exe"
SBECMD_PATH = "SBECmd\\SBECmd.exe"
PECMD_PATH = "PECmd\\PECmd.exe"

LECMD_INPUT_FILE = None # folder for multiple files
SBECMD_INPUT_FILE = None # single file
SBECMD_INPUT_DIR = os.path.dirname(os.path.abspath(__file__))
PECMD_INPUT_FILE = None # single file

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

LECMD_CSV_PATH = OUTPUT_DIR
SBECMD_CSV_PATH = OUTPUT_DIR
PECMD_CSV_PATH = OUTPUT_DIR
FINAL_CSV_PATH = OUTPUT_DIR

LECMD_OUT_FILE_NAME = "link_output.csv"
RECMD_OUT_FILE_NAME = "shellb_output.csv"
PECMD_OUT_FILE_NAME = "prefetch_output.csv"
PECMD_OUT_FILE_NAME2 = "prefetch_output_Timeline.csv"
FINAL_OUT_FILE_NAME = "timeline_output_merged.xlsx"

def check_path_exists(path):
    if os.path.exists(path):
        if os.path.isfile(path):
            return path
        elif os.path.isdir(path):
            return path
    return "Not Found"

def getFileInputDirectory(lnk_dir, sbe_dir, pf_dir):
    
    print("\n--- Detected Directory Files ---")
    print("LNK dir: " + check_path_exists(lnk_dir))
    print("SBE dir: " + check_path_exists(sbe_dir))
    print("PF dir: " + check_path_exists(pf_dir))
    print("---       END      ---\n")

def getFileInputNames():
    directory = os.path.dirname(os.path.abspath(__file__))

    pref_file = None
    ln_file = None
    shell_file = None

    for file in os.listdir(directory):
        if os.path.isfile(os.path.join(directory, file)):  # Ensure it's a file, not a folder
            if file.endswith(".lnk"):
                ln_file = OUTPUT_DIR
            elif file.endswith(".pf"):
                pref_file = file
            elif "." not in file:
                shell_file = file

    print("\n--- Detected Files ---")
    print("LINK File:", OUTPUT_DIR if ln_file else "Not Found")
    print("USR File:", shell_file if shell_file else "Not Found")
    print("PREF File:", pref_file if pref_file else "Not Found")
    print("---       END      ---\n")

    return ln_file, shell_file, pref_file

# Function to run RECmd and process registry data
def run_sbecmd(SBECMD_INPUT_FILE):
    sbecmd_command = SBECMD_PATH + " -d " + SBECMD_INPUT_DIR + " --csv " + SBECMD_CSV_PATH + " --csvf " + RECMD_OUT_FILE_NAME

    try:
        print(f"Running SBECmd on {SBECMD_INPUT_FILE}...")
        subprocess.run(sbecmd_command, shell=True, check=True)
        print(f"Shellbags parsed successfully! Output saved to: {SBECMD_CSV_PATH}")
    except subprocess.CalledProcessError as e:
        print(f"Error running RECmd: {e}")

# Function to run AmcacheParser and process execution history
def run_pecmd(PECMD_INPUT_FILE):
    pecmd_command = f'"{PECMD_PATH}" -f "{PECMD_INPUT_FILE}" --csv "{PECMD_CSV_PATH}" --csvf "{PECMD_OUT_FILE_NAME}"'

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
        print(f"Link directory parsed successfully! Output saved to: {LECMD_CSV_PATH}")
    except subprocess.CalledProcessError as e:
        print(f"Error running LECmd: {e}")

def get_latest_shellb_output_file(directory="."):
    latest_file = None
    latest_mtime = 0

    for file in os.listdir(directory):
        file_path = os.path.join(directory, file)
        if "shellb_output.csv" in file and os.path.isfile(file_path):
            file_mtime = os.path.getmtime(file_path)
            if file_mtime > latest_mtime:
                latest_file = file_path
                latest_mtime = file_mtime

    if latest_file:
        print(f"Using latest file: {latest_file}")
    else:
        print("No 'shellb_output.csv' files found.")

    return latest_file

def adjust_column_widths(writer, df, sheet_name):
    worksheet = writer.sheets[sheet_name]
    for i, col in enumerate(df.columns):
        max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, max_len)

def runAllTools(LECMD_INPUT_FILE, PECMD_INPUT_FILE, SBECMD_INPUT_FILE):
    run_lecmd(LECMD_INPUT_FILE)
    run_sbecmd(SBECMD_INPUT_FILE)
    run_pecmd(PECMD_INPUT_FILE)


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
        LECMD_INPUT_FILE, SBECMD_INPUT_FILE,  PECMD_INPUT_FILE = getFileInputNames()
        print(LECMD_INPUT_FILE)
        time.sleep(1)
        runAllTools(LECMD_INPUT_FILE, PECMD_INPUT_FILE, SBECMD_INPUT_FILE)
    
    if val == "2":
        print("\nNOTE: Ensure that LIVE files are not scanned as the EvtxECmd Tool does not allow it.")
        print("Also, write the directory name with no quotation marks (\"\") or double backslashes (\\\\).\n")
        
        print("--- Specify User Directory ---")

        # "C:\Windows\System32\winevt\Logs\Internet Explorer.evtx" - sample dir
        sb_dir = input("SB Directory: ")
        # "C:\Windows\Prefetch\WSL.EXE-1B26B53B.pf" - sample dir
        pf_dir = input("PF Directory: ")
        # "C:\Windows\System32\config\SOFTWARE" - sample dir
        lnk_dir = input("LNK Directory: ")

        print("Scanning directory for specific files...")
        time.sleep(1)
        getFileInputDirectory(sb_dir, pf_dir, lnk_dir)
        time.sleep(1)

        runAllTools(lnk_dir, pf_dir, sb_dir)

    if val == "3":
        print("Exiting...")
        os._exit(0);

if __name__ == "__main__":
    main()