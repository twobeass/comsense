"""
list_com_progids.py: List all registered COM ProgIDs available for extraction

Usage:
    python list_com_progids.py

Description:
    Scans the Windows registry (HKEY_CLASSES_ROOT) and prints all ProgIDs which have a CLSID subkey.
    This helps users determine which names to use as input for extract_com.py.
"""
import winreg


def list_progids():
    progids = []
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "") as root:
        i = 0
        while True:
            try:
                subkey_name = winreg.EnumKey(root, i)
                if '.' in subkey_name:
                    try:
                        with winreg.OpenKey(root, f"{subkey_name}\\CLSID") as clsid_key:
                            clsid, _ = winreg.QueryValueEx(clsid_key, "")
                            progids.append((subkey_name, clsid))
                    except FileNotFoundError:
                        pass
                i += 1
            except OSError:
                break
    return progids


def main():
    print("Listing available COM ProgIDs (names to use with extract_com.py):\n")
    print(f"{'ProgID':40} | {'CLSID'}")
    print("-" * 80)
    for progid, clsid in sorted(list_progids()):
        print(f"{progid:40} | {clsid}")


if __name__ == "__main__":
    main()
