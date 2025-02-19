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
REGISTRY_INPUT_FILE = "C:\\Windows\\System32\\config\\SOFTWARE"
REGISTRY_BATCH_FILE = "RECmd\\BatchExamples\\SoftwareASEPs.reb"
PREFETCH_INPUT_FILE = "XMIND.EXE-A9C728F9.pf"

OUTPUT_DIR = ".\\"

EVTX_CSV_PATH = os.path.join(OUTPUT_DIR, "zOuput")
REGISTRY_CSV_PATH = os.path.join(OUTPUT_DIR, "zOuput")
PREFETCH_CSV_PATH = os.path.join(OUTPUT_DIR, "zOuput")
FINAL_CSV_PATH = os.path.join(OUTPUT_DIR, "zOuput")

os.makedirs(OUTPUT_DIR, exist_ok=True)

def run_evtxecmd():
    command = f'"{EVTXECMD_PATH}" -f "{EVTX_INPUT_FILE}" --sd "{day_scope}" --csv "{EVTX_CSV_PATH}"' #--maps "{EVTXECMD_PATH}\\Maps"'
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
    recmd_command = f'"{RECMD_PATH}" -f "{REGISTRY_INPUT_FILE}" --bn "{REGISTRY_BATCH_FILE}" --csv "{REGISTRY_CSV_PATH}"'  # --sk parses all known keys

    try:
        print(f"Running RECmd on {REGISTRY_INPUT_FILE}...")
        subprocess.run(recmd_command, shell=True, check=True)
        print(f"Registry parsed successfully! Output saved to: {REGISTRY_CSV_PATH}")
    except subprocess.CalledProcessError as e:
        print(f"Error running RECmd: {e}")

def run_pecmd():
    command = f'"{PECMD_PATH}" -f "{PREFETCH_INPUT_FILE}" --csv "{PREFETCH_CSV_PATH}"'

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

if __name__ == "__main__":
    run_evtxecmd()
    run_recmd()
    run_pecmd()