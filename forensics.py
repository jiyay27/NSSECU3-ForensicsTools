import os
import time
import pandas as pd
import subprocess
from datetime import datetime, timedelta, timezone
utc_now = datetime.now(timezone.utc)
day_scope = utc_now - timedelta(days=30)
day_scope = day_scope.strftime("%Y-%m-%d %H:%M:%S")

EVTXECMD_PATH = "EvtxECmd\\EvtxECmd.exe"
RECMD_PATH = "RECmd\\RECmd.exe"
PECMD_PATH = "PECmd\\PECmd.exe"

EVTX_INPUT_FILE = "OAlerts.evtx"
REGISTRY_INPUT_FILE = "SOFTWARE"
REGISTRY_BATCH_FILE = "RECmd\\BatchExamples\\SoftwareASEPs.reb"
PREFETCH_INPUT_FILE = "XMIND.EXE-A9C728F9.pf"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

OUTPUT_DIR = SCRIPT_DIR

EVTX_CSV_PATH = OUTPUT_DIR
REGISTRY_CSV_PATH = OUTPUT_DIR
PREFETCH_CSV_PATH = OUTPUT_DIR
FINAL_CSV_PATH = OUTPUT_DIR

EVTX_OUT_FILE_NAME = "evtx_output.csv"
RECMD_OUT_FILE_NAME = "registry_output.csv"
PECMD_OUT_FILE_NAME = "prefetch_output.csv"
PECMD2_OUT_FILE_NAME = "prefetch_output_Timeline.csv"
FINAL_OUT_FILE_NAME = "forensic_output_merged.xlsx"


os.makedirs(OUTPUT_DIR, exist_ok=True)

def run_evtxecmd():
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


def run_recmd():
    recmd_command = f'"{RECMD_PATH}" -f "{REGISTRY_INPUT_FILE}" --bn "{REGISTRY_BATCH_FILE}" --csv {REGISTRY_CSV_PATH} --csvf {RECMD_OUT_FILE_NAME}'

    try:
        print(f"Running RECmd on {REGISTRY_INPUT_FILE}...")
        subprocess.run(recmd_command, shell=True, check=True)
        print(f"Registry parsed successfully! Output saved to: {REGISTRY_CSV_PATH}")
    except subprocess.CalledProcessError as e:
        print(f"Error running RECmd: {e}")

def run_pecmd():
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

    evtx_df = pd.read_csv(EVTX_OUT_FILE_NAME, low_memory=False) if os.path.exists(EVTX_CSV_PATH) else pd.DataFrame()
    registry_df = pd.read_csv(RECMD_OUT_FILE_NAME, low_memory=False) if os.path.exists(REGISTRY_CSV_PATH) else pd.DataFrame()
    prefetch_df = pd.read_csv(PECMD_OUT_FILE_NAME, low_memory=False) if os.path.exists(PREFETCH_CSV_PATH) else pd.DataFrame()
    prefetch2_df = pd.read_csv(PECMD2_OUT_FILE_NAME, low_memory=False) if os.path.exists(PREFETCH_CSV_PATH) else pd.DataFrame()

    if os.path.exists(FINAL_OUT_FILE_NAME):
        print(f"\nFile '{FINAL_OUT_FILE_NAME}' already exists. Overwriting...")
        os.remove(FINAL_OUT_FILE_NAME)

    time.sleep(1)

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

    if os.path.exists(FINAL_OUT_FILE_NAME):
        print(f"Forensic timeline created successfully! Saved at: {FINAL_CSV_PATH}")
    else:
        print("Error creating forensic timeline")

    os.remove(EVTX_OUT_FILE_NAME)
    os.remove(RECMD_OUT_FILE_NAME)
    os.remove(PECMD_OUT_FILE_NAME)
    os.remove(PECMD2_OUT_FILE_NAME)


if __name__ == "__main__":
    # print("Current Working Directory:", os.getcwd())
     run_evtxecmd()
     run_recmd()
     run_pecmd()
     
     create_forensic_timeline()