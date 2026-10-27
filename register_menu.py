import os
import sys
import winreg
import win32com.client

def install_integration():
    python_exe = sys.executable
    pythonw_exe = python_exe.replace("python.exe", "pythonw.exe")
    if not os.path.exists(pythonw_exe):
        pythonw_exe = python_exe

    # Always find ui.py in the exact same folder as register_menu.py
    script_dir = os.path.dirname(os.path.abspath(__file__))
    ui_script = os.path.join(script_dir, "ui.py")
    
    if not os.path.exists(ui_script):
        print(f"Error: Could not find ui.py at {ui_script}")
        return

    # ---------------------------------------------------------
    # HOOK 1: Direct Right-Click Menu (Windows Registry)
    # ---------------------------------------------------------
    key_path = r"Software\Classes\*\shell\QuickConvert"
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
        winreg.SetValue(key, "", winreg.REG_SZ, "QuickConvert")
        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "shell32.dll,238")
        winreg.CloseKey(key)

        cmd_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, f"{key_path}\\command")
        command_str = f'"{pythonw_exe}" "{ui_script}" "%1"'
        winreg.SetValue(cmd_key, "", winreg.REG_SZ, command_str)
        winreg.CloseKey(cmd_key)
        print("Added 'QuickConvert' to the main Windows Right-Click Menu!")
    except Exception as e:
        print(f"Registry error: {e}")

    # ---------------------------------------------------------
    # HOOK 2: Multi-File Batch Hook ('Send to' Menu)
    # ---------------------------------------------------------
    try:
        sendto_dir = os.path.join(
            os.environ["APPDATA"], "Microsoft", "Windows", "SendTo"
        )
        # Keep filename ASCII-clean for WScript.Shell compatibility
        shortcut_path = os.path.join(sendto_dir, "QuickConvert (Multi-File).lnk")

        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(shortcut_path)
        shortcut.TargetPath = pythonw_exe
        shortcut.Arguments = f'"{ui_script}"'
        shortcut.WorkingDirectory = script_dir
        shortcut.IconLocation = "shell32.dll,238"
        shortcut.Save()
        print("Added 'QuickConvert (Multi-File)' to the Right-Click -> 'Send to' Menu!")
    except Exception as e:
        print(f"SendTo error: {e}")


def uninstall_integration():
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\*\shell\QuickConvert\command")
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\*\shell\QuickConvert")
        print("Removed from main Right-Click Menu.")
    except FileNotFoundError:
        pass

    sendto_lnk = os.path.join(
        os.environ["APPDATA"], "Microsoft", "Windows", "SendTo", "QuickConvert (Multi-File).lnk"
    )
    if os.path.exists(sendto_lnk):
        os.remove(sendto_lnk)
        print("Removed from 'Send to' Menu.")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1].lower() == "remove":
        uninstall_integration()
    else:
        install_integration() 