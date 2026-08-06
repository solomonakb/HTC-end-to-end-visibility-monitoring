"""
Run this directly on the machine that hosts waitress_server.py, under the
SAME account the app actually runs as (e.g. if it runs as a Windows Service,
run this via that service's "Log On As" account, or via psexec -u <svc_account>).

    python diagnose_share.py

This tells you exactly which of these is failing:
  1. DNS / name resolution of the server (svhqftp03)
  2. SMB port reachability (445)
  3. Directory listing (os.listdir) - permissions / credentials
  4. Extension matching (in case files ARE visible but named unexpectedly)
"""
import os
import socket
import getpass

SHARE_PATH = r'\\svhqftp03\HCT'
SERVER = 'svhqftp03'

print(f"Running as user: {getpass.getuser()}")
print(f"Target share:    {SHARE_PATH}\n")

# 1. DNS resolution
print("[1/4] Resolving server name...")
try:
    ip = socket.gethostbyname(SERVER)
    print(f"      OK -> {SERVER} resolves to {ip}")
except socket.gaierror as e:
    print(f"      FAIL -> Could not resolve '{SERVER}': {e}")
    print("      This machine's DNS/hosts file doesn't know this server.")
    print("      (If it works fine in Explorer on your own PC, the service")
    print("       host may be on a different VLAN/DNS scope than your PC.)")

# 2. Port 445 (SMB) reachability
print("\n[2/4] Testing SMB port (445) reachability...")
try:
    with socket.create_connection((SERVER, 445), timeout=5):
        print(f"      OK -> Port 445 open on {SERVER}")
except Exception as e:
    print(f"      FAIL -> Could not reach {SERVER}:445 - {e}")
    print("      Firewall, VPN, or the server being on a different network")
    print("      segment than wherever this script is running are the")
    print("      usual causes.")

# 3. Actual directory listing (the real test - matches what the app does)
print(f"\n[3/4] Listing '{SHARE_PATH}'...")
try:
    entries = os.listdir(SHARE_PATH)
    print(f"      OK -> {len(entries)} item(s) found")
except FileNotFoundError as e:
    print(f"      FAIL -> Path not found: {e}")
except PermissionError as e:
    print(f"      FAIL -> Access denied: {e}")
    print("      The account running THIS process has no credentials for")
    print("      the share. This is the #1 cause when 'it works when I")
    print("      browse it myself' but fails from the running service.")
    print("      Fix: run the service under a domain account that has")
    print("      access, or map credentials for the service account, e.g.")
    print("      from an elevated prompt running AS that service account:")
    print(r'          net use \\svhqftp03\HCT /user:DOMAIN\svc_account *')
    entries = []
except OSError as e:
    print(f"      FAIL -> OS error (winerror={getattr(e, 'winerror', None)}): {e}")
    entries = []

# 4. What extensions are actually there
if entries:
    print(f"\n[4/4] Sample of what's in the folder (first 20):")
    for name in entries[:20]:
        print(f"      {name}")
    xlsx_like = [e for e in entries if e.lower().endswith(('.xlsx', '.xls'))]
    print(f"\n      {len(xlsx_like)} of {len(entries)} match .xlsx/.xls")
    if not xlsx_like and entries:
        print("      None matched .xlsx/.xls - check actual extensions above")
        print("      (e.g. .xlsm, .csv, or files inside subfolders).")
else:
    print("\n[4/4] Skipped (folder listing failed above).")