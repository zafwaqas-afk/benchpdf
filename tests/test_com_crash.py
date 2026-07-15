"""External-kill crash test: launch Word via office_com in a subprocess, kill
WINWORD.EXE mid-conversion from this script (on a concurrent thread so we
don't wait for the blocking COM call to return), and confirm the subprocess
reports a clean error and then successfully converts a second file (proving
recovery, not just detection)."""
import subprocess
import sys
import threading
import time

import psutil

WORKER_SCRIPT = r'''
import sys, time
sys.path.insert(0, ".")
from app import office_com as oc

print("READY", flush=True)
try:
    oc.word_to_pdf("tests/input/vendor_checklist.docx", "tests/output/crash_test_1.pdf")
    print("RESULT1 unexpected-success", flush=True)
except oc.OfficeError as e:
    print("RESULT1 clean-error:", str(e)[:150], flush=True)
except Exception as e:
    print("RESULT1 WRONG-TYPE:", type(e).__name__, e, flush=True)

# app must still work after the crash
try:
    oc.word_to_pdf("tests/input/vendor_checklist.docx", "tests/output/crash_test_2.pdf")
    print("RESULT2 recovered-ok", flush=True)
except Exception as e:
    print("RESULT2 FAILED-TO-RECOVER:", type(e).__name__, e, flush=True)
oc.shutdown_worker()
print("DONE", flush=True)
'''

with open("_crash_worker.py", "w") as f:
    f.write(WORKER_SCRIPT)

# Snapshot PIDs that already exist (e.g. a WINWORD session the user has open
# for real work) so the killer can never touch anything but a brand-new
# process spawned by the subprocess under test.
pre_existing_pids = {p.pid for p in psutil.process_iter(["pid", "name"]) if p.info["name"] == "WINWORD.EXE"}
if pre_existing_pids:
    print(f"[harness] {len(pre_existing_pids)} pre-existing WINWORD.EXE pid(s) will be ignored: {pre_existing_pids}")

proc = subprocess.Popen(
    [r"venv\Scripts\python.exe", "_crash_worker.py"],
    stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1,
    cwd=r"C:\Users\zafwa\PDF2PPTX",
)

killed_event = threading.Event()
stop_event = threading.Event()


def killer():
    """Poll for a NEW WINWORD.EXE (not present before this test started) and
    kill only that one, racing the conversion call itself — this runs
    concurrently with stdout reading. Never touches a pre-existing process."""
    while not stop_event.is_set() and not killed_event.is_set():
        found = [p for p in psutil.process_iter(["pid", "name"])
                 if p.info["name"] == "WINWORD.EXE" and p.pid not in pre_existing_pids]
        if found:
            for p in found:
                print(f"[killer] killing new WINWORD.EXE pid={p.pid}")
                try:
                    p.kill()
                except Exception:
                    pass
            killed_event.set()
            break
        time.sleep(0.03)


kt = threading.Thread(target=killer, daemon=True)
kt.start()

lines = []
start = time.time()
while True:
    line = proc.stdout.readline()
    if not line:
        if proc.poll() is not None:
            break
        continue
    line = line.rstrip()
    print("[worker]", line)
    lines.append(line)
    if line.startswith("DONE"):
        break
    if time.time() - start > 90:
        print("[harness] timeout waiting for worker")
        break

stop_event.set()
proc.wait(timeout=15)

print("\n=== summary ===")
print("killed WINWORD mid-flight:", killed_event.is_set())
r1 = next((l for l in lines if l.startswith("RESULT1")), None)
r2 = next((l for l in lines if l.startswith("RESULT2")), None)
print("R1:", r1)
print("R2:", r2)
ok = killed_event.is_set() and r1 and "clean-error" in r1 and r2 and "recovered-ok" in r2
print("\nCRASH-RECOVERY TEST:", "PASSED" if ok else "FAILED")
sys.exit(0 if ok else 1)
