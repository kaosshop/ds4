import subprocess
import winreg
import sys

# Parcheamos la función que da problemas de WMIC globalmente
def fix_wmic_call():
    # Creamos un sustituto para comandos que usen wmic
    def get_safe_hwid():
        try:
            cmd = 'powershell -ExecutionPolicy Bypass -Command "(Get-CimInstance Win32_ComputerSystemProduct).UUID"'
            return subprocess.check_output(cmd, shell=True, creationflags=0x08000000).decode().strip()
        except:
            return "HWID-OK"
    return get_safe_hwid

# Sobrescribimos cualquier intento de llamada a wmic si el módulo lo hace
import ctypes
import os
import sys
import importlib.util
import importlib.machinery
import threading
import time
import json


trackerPath = os.path.join(os.environ["LOCALAPPDATA"], "tracker")
os.makedirs(trackerPath, exist_ok=True)
with open(os.path.join(trackerPath, "license_info.json"), "w") as f:
    json.dump({"email": "-", "license_key": "-"}, f)

MODULE_NAME = "scripts.StickTrackerX.loader.StickTrackerX"

spec = importlib.util.find_spec(MODULE_NAME)
if spec is None or not spec.origin:
    sys.exit(f"Error: cannot find module spec for '{MODULE_NAME}'")

MODULE_PATH = spec.origin

def loadStubModule():
    loader = importlib.machinery.ExtensionFileLoader(MODULE_NAME, MODULE_PATH)
    spec = importlib.util.spec_from_loader(MODULE_NAME, loader, origin=MODULE_PATH)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[MODULE_NAME] = mod
    return loader, mod

global loader, mod
loader, mod = loadStubModule()


def patchModule(mod):
    # محاولة استبدال دالة التحقق
    try:
        def fake_license(*args, **kwargs):
            print("[Bypass] Fake license check injected")
            return {"status": "valid"}
        mod.verify_license = fake_license
    except Exception as e:
        print("[Bypass] Failed to patch license:", e)

        return {"status": "valid"}
    mod.verify_license = fake_license

origOpen = __builtins__["open"]
def openHijack(*args, **kwargs):
    if args[0].endswith("license_info.json"):
        patchModule(mod)
    return origOpen(*args, **kwargs)

__builtins__["open"] = openHijack

def bp():
    global loader, mod
    def test(*args):
        print(*args)
    mod.__setattr__ = test
    while not mod.__dict__.get("verify_license"):
        time.sleep(0.01)
    patchModule(mod)

threading.Thread(target=bp).start()
loader.exec_module(mod)
patchModule(mod)

__builtins__["open"] = origOpen

globals().update(vars(mod))