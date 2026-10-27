import sys
import os
import winreg
import win32gui
import win32com.client
import keyboard
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QHBoxLayout, QVBoxLayout,
    QPushButton, QComboBox, QCheckBox, QSystemTrayIcon, QMenu,
    QStyle, QFileDialog, QDialog, QKeySequenceEdit
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSettings, QObject
from PyQt6.QtGui import QAction, QKeySequence

from engine import convert_files


# =====================================================================
# 1. BUILT-IN EXTENSION INSTALLER (Startup + Right-Click + SendTo)
# =====================================================================
def install_as_extension():
    """Registers the compiled .exe to run on Windows Startup and Right-Click menus."""
    if getattr(sys, 'frozen', False):
        exe_path = os.path.abspath(sys.executable)
        work_dir = os.path.dirname(exe_path)
        cmd_base = f'"{exe_path}"'
    else:
        pythonw = sys.executable.replace("python.exe", "pythonw.exe")
        script_path = os.path.abspath(__file__)
        work_dir = os.path.dirname(script_path)
        exe_path = pythonw
        cmd_base = f'"{pythonw}" "{script_path}"'

    # A. Add to Windows Startup
    run_key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key_path, 0, winreg.KEY_SET_VALUE) as key:
        winreg.SetValueEx(key, "QuickConvert", 0, winreg.REG_SZ, f"{cmd_base} --tray")

    # B. Add to Right-Click Context Menu
    ctx_path = r"Software\Classes\*\shell\QuickConvert"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, ctx_path) as key:
        winreg.SetValue(key, "", winreg.REG_SZ, "QuickConvert")
        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "shell32.dll,238")
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, f"{ctx_path}\\command") as cmd_key:
        winreg.SetValue(cmd_key, "", winreg.REG_SZ, f'{cmd_base} "%1"')

    # C. Add to 'Send to' Menu (For multi-file selection)
    sendto_dir = os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "SendTo")
    lnk_path = os.path.join(sendto_dir, "QuickConvert (Multi-File).lnk")
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(lnk_path)
    shortcut.TargetPath = exe_path
    if not getattr(sys, 'frozen', False):
        shortcut.Arguments = f'"{os.path.abspath(__file__)}"'
    shortcut.WorkingDirectory = work_dir
    shortcut.IconLocation = "shell32.dll,238"
    shortcut.Save()


def uninstall_extension():
    """Removes Startup, Right-Click, and SendTo hooks."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE) as key:
            winreg.DeleteValue(key, "QuickConvert")
    except FileNotFoundError:
        pass

    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\*\shell\QuickConvert\command")
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\*\shell\QuickConvert")
    except FileNotFoundError:
        pass

    lnk_path = os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "SendTo", "QuickConvert (Multi-File).lnk")
    if os.path.exists(lnk_path):
        os.remove(lnk_path)


# =====================================================================
# 2. FILE EXPLORER SELECTION READER & WORKER THREAD
# =====================================================================
def get_explorer_selection() -> list[str]:
    """Reads selected files from the active File Explorer window/tab."""
    shell = win32com.client.Dispatch("Shell.Application")
    foreground_hwnd = win32gui.GetForegroundWindow()
    fallback_paths = []

    for window in shell.Windows():
        try:
            items = window.Document.SelectedItems()
            if items.Count > 0:
                paths = [items.Item(i).Path for i in range(items.Count)]
                if window.HWND == foreground_hwnd and len(paths) > 1:
                    return paths
                if len(paths) > len(fallback_paths):
                    fallback_paths = paths
        except Exception:
            continue
    return fallback_paths


class ConversionWorker(QThread):
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, files, target_format, merge_files, custom_output):
        super().__init__()
        self.files = files
        self.target_format = target_format
        self.merge_files = merge_files
        self.custom_output = custom_output

    def run(self):
        try:
            outputs = convert_files(
                self.files, self.target_format, self.merge_files, self.custom_output
            )
            self.finished.emit(outputs)
        except Exception as e:
            self.error.emit(str(e))


# =====================================================================
# 3. GLOBAL HOTKEY BRIDGE & RECORDER DIALOG
# =====================================================================
class HotkeyBridge(QObject):
    triggered = pyqtSignal()


class ShortcutDialog(QDialog):
    def __init__(self, current_hotkey: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Set Custom Shortcut")
        self.setFixedSize(300, 140)
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Dialog)
        self.setStyleSheet("""
            QDialog { background-color: #202020; color: #FFFFFF; font-family: 'Segoe UI'; }
            QLabel { font-size: 12px; color: #CCCCCC; border: none; }
            QKeySequenceEdit {
                background-color: #2D2D2D; color: #60CDFF;
                border: 1px solid #0078D4; border-radius: 4px;
                padding: 6px; font-size: 13px; font-weight: bold;
            }
            QPushButton {
                background-color: #0078D4; color: white;
                border: none; border-radius: 4px; padding: 6px 12px; font-weight: bold;
            }
            QPushButton:hover { background-color: #1084E0; }
        """)

        layout = QVBoxLayout(self)
        lbl = QLabel("Click the box below and press your desired\nkeyboard shortcut (e.g., Alt+C or Ctrl+Shift+Q):")
        layout.addWidget(lbl)

        self.key_editor = QKeySequenceEdit(QKeySequence(current_hotkey))
        layout.addWidget(self.key_editor)

        btn_layout = QHBoxLayout()
        save_btn = QPushButton("Save Shortcut")
        save_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setStyleSheet("background-color: #3D3D3D;")
        cancel_btn.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(save_btn)
        layout.addLayout(btn_layout)

    def get_shortcut(self) -> str:
        return self.key_editor.keySequence().toString().lower()


# =====================================================================
# 4. TOP-BAR UI + SYSTEM TRAY APPLICATION
# =====================================================================
class TopBarConverter(QWidget):
    def __init__(self, preloaded_files=None):
        super().__init__()
        self.preloaded_files = preloaded_files or []
        
        self.settings = QSettings("QuickConvert", "UserPreferences")
        self.current_hotkey = self.settings.value("hotkey", "alt+c")

        self.hotkey_bridge = HotkeyBridge()
        self.hotkey_bridge.triggered.connect(self.on_hotkey_pressed)

        self.init_ui()
        self.init_tray()
        self.register_global_hotkey(self.current_hotkey)

        if self.preloaded_files:
            self.status_lbl.setText(f"📂 {len(self.preloaded_files)} ready")

    def init_ui(self):
        self.setWindowTitle("QuickConvert")
        self.setFixedSize(650, 44)
        self.setWindowFlags(
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.Tool
        )
        self.setAcceptDrops(True)

        screen = QApplication.primaryScreen().geometry()
        self.move((screen.width() - self.width()) // 2 + 80, 48)

        self.setStyleSheet("""
            QWidget {
                background-color: #202020; color: #FFFFFF;
                font-family: 'Segoe UI'; font-size: 12px;
                border: 1px solid #383838; border-radius: 6px;
            }
            QPushButton {
                background-color: #2D2D2D; border: 1px solid #404040;
                border-radius: 4px; padding: 4px 10px; font-weight: 600;
            }
            QPushButton:hover { background-color: #3D3D3D; }
            QComboBox {
                background-color: #2D2D2D; border: 1px solid #404040;
                border-radius: 4px; padding: 3px 8px;
            }
        """)

        layout = QHBoxLayout()
        layout.setContentsMargins(10, 4, 10, 4)
        layout.setSpacing(8)

        self.status_lbl = QLabel(f"⚡ QuickConvert ({self.current_hotkey.upper()})")
        self.status_lbl.setStyleSheet("border: none; font-weight: bold; color: #60CDFF;")
        layout.addWidget(self.status_lbl)

        self.fmt_combo = QComboBox()
        self.fmt_combo.addItems(["PDF", "DOCX", "PNG", "JPG", "XLSX", "CSV"])
        layout.addWidget(self.fmt_combo)

        self.merge_cb = QCheckBox("Merge Files")
        self.merge_cb.setChecked(True)
        self.merge_cb.setStyleSheet("border: none;")
        layout.addWidget(self.merge_cb)

        self.convert_btn = QPushButton("Convert Selected")
        self.convert_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.convert_btn.setStyleSheet("background-color: #0078D4; color: white; border: none;")
        self.convert_btn.clicked.connect(self.convert_highlighted_files)
        layout.addWidget(self.convert_btn)

        # Manual Multi-File Picker Button (100% immune to Win11 Explorer Tab bugs)
        self.pick_btn = QPushButton("📂 Pick")
        self.pick_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.pick_btn.setToolTip("Select multiple files from anywhere to merge/convert")
        self.pick_btn.clicked.connect(self.pick_files_manually)
        layout.addWidget(self.pick_btn)

        # Shortcut Settings Button
        hk_btn = QPushButton("⌨️")
        hk_btn.setFixedWidth(30)
        hk_btn.setToolTip(f"Change Keyboard Shortcut (Current: {self.current_hotkey.upper()})")
        hk_btn.clicked.connect(self.prompt_custom_shortcut)
        layout.addWidget(hk_btn)

        # Hide button
        hide_btn = QPushButton("—")
        hide_btn.setFixedWidth(26)
        hide_btn.setToolTip("Minimize to System Tray")
        hide_btn.setStyleSheet("border: none; background: transparent; color: #999;")
        hide_btn.clicked.connect(self.hide)
        layout.addWidget(hide_btn)

        self.setLayout(layout)

    def register_global_hotkey(self, hotkey_str: str):
        try:
            keyboard.unhook_all_hotkeys()
        except Exception:
            pass

        if not hotkey_str:
            return

        try:
            keyboard.add_hotkey(hotkey_str, lambda: self.hotkey_bridge.triggered.emit())
            self.current_hotkey = hotkey_str
            self.settings.setValue("hotkey", hotkey_str)
            self.status_lbl.setText(f"⚡ QuickConvert ({hotkey_str.upper()})")
        except Exception:
            self.status_lbl.setText("⚠️ Invalid Hotkey")

    def prompt_custom_shortcut(self):
        dlg = ShortcutDialog(self.current_hotkey, self)
        if dlg.exec():
            new_hk = dlg.get_shortcut()
            if new_hk:
                self.register_global_hotkey(new_hk)

    def on_hotkey_pressed(self):
        self.show()
        self.raise_()
        self.activateWindow()
        self.convert_highlighted_files()

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        icon = self.style().standardIcon(QStyle.StandardPixmap.SP_ArrowRight)
        self.tray.setIcon(icon)
        self.tray.setToolTip("QuickConvert (Click to toggle Top Bar)")

        menu = QMenu()
        toggle_action = QAction("Show / Hide Top Bar", self)
        toggle_action.triggered.connect(self.toggle_visibility)

        shortcut_action = QAction("Set Custom Keyboard Shortcut...", self)
        shortcut_action.triggered.connect(self.prompt_custom_shortcut)

        install_action = QAction("Enable Auto-Start & Right-Click", self)
        install_action.triggered.connect(install_as_extension)

        uninstall_action = QAction("Remove Extension Hooks", self)
        uninstall_action.triggered.connect(uninstall_extension)

        quit_action = QAction("Quit Completely", self)
        quit_action.triggered.connect(QApplication.instance().quit)

        menu.addAction(toggle_action)
        menu.addAction(shortcut_action)
        menu.addSeparator()
        menu.addAction(install_action)
        menu.addAction(uninstall_action)
        menu.addSeparator()
        menu.addAction(quit_action)

        self.tray.setContextMenu(menu)
        self.tray.activated.connect(self.on_tray_click)
        self.tray.show()

    def toggle_visibility(self):
        if self.isVisible():
            self.hide()
        else:
            self.show()
            self.raise_()

    def on_tray_click(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self.toggle_visibility()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.MouseButton.LeftButton:
            self.move(event.globalPosition().toPoint() - self.drag_pos)

    def convert_highlighted_files(self):
        files = self.preloaded_files if self.preloaded_files else get_explorer_selection()
        if not files:
            self.status_lbl.setText("⚠️ Select file(s) first!")
            return
        self.run_engine(files)

    def pick_files_manually(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files to Convert or Merge")
        if files:
            self.run_engine(files)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls() if os.path.isfile(u.toLocalFile())]
        if files:
            self.run_engine(files)

    def run_engine(self, files):
        target = self.fmt_combo.currentText().lower()
        merge = self.merge_cb.isChecked()
        first_dir = os.path.dirname(os.path.abspath(files[0]))
        first_base = os.path.splitext(os.path.basename(files[0]))[0]

        is_single_output = (len(files) == 1) or (merge and target in ["pdf", "docx", "xlsx", "csv"])

        if is_single_output:
            default_name = f"merged_output.{target}" if len(files) > 1 else f"{first_base}_converted.{target}"
            default_path = os.path.join(first_dir, default_name)

            custom_output, _ = QFileDialog.getSaveFileName(
                self,
                f"Save {len(files)} File(s) As",
                default_path,
                f"{target.upper()} Files (*.{target});;All Files (*.*)"
            )
        else:
            custom_output = QFileDialog.getExistingDirectory(
                self,
                "Select Destination Folder for Converted Files",
                first_dir
            )

        if not custom_output:
            self.status_lbl.setText(f"⚡ QuickConvert ({self.current_hotkey.upper()})")
            return

        self.status_lbl.setText(f"⏳ Converting {len(files)}...")
        self.convert_btn.setEnabled(False)

        self.worker = ConversionWorker(files, target, merge, custom_output)
        self.worker.finished.connect(self.on_done)
        self.worker.error.connect(self.on_err)
        self.worker.start()

    def on_done(self, outputs):
        self.convert_btn.setEnabled(True)
        self.preloaded_files = []
        out_name = os.path.basename(outputs[0]) if len(outputs) == 1 else f"{len(outputs)} files"
        self.status_lbl.setText(f"✅ Saved: {out_name[:14]}")

    def on_err(self, err):
        self.convert_btn.setEnabled(True)
        print(f"ERROR: {err}")
        self.status_lbl.setText(f"❌ {str(err)[:20]}")
        self.status_lbl.setToolTip(str(err))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    install_as_extension()

    cli_files = [os.path.abspath(f) for f in sys.argv[1:] if not f.startswith("--") and os.path.exists(f)]
    bar = TopBarConverter(preloaded_files=cli_files)
    bar.show()
    sys.exit(app.exec())