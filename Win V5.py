import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
from plyer import notification
import subprocess
import datetime
import threading
import os
import sys
import ctypes
import socket
import json

LOG_FILE = "update_log.txt"

# === Admin Check ===
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def notify_user(title, message):
    notification.notify(title=title, message=message, timeout=10)

def log_message(message):
    timestamp = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S] ")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(timestamp + message + "\n")

def run_powershell(command):
    try:
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-Command", command],
            capture_output=True,
            text=True
        )
        return result.stdout.strip(), result.stderr.strip()
    except Exception as e:
        return "", str(e)

def has_internet():
    try:
        socket.create_connection(("8.8.8.8", 53), timeout=3)
        return True
    except OSError:
        return False

def fetch_updates():
    command = (
        "Import-Module PSWindowsUpdate; "
        "Get-WindowsUpdate | Select-Object -Property KB, Title, Size, Description | ConvertTo-Json"
    )
    stdout, stderr = run_powershell(command)
    if stderr:
        log_message("[Fehler beim Abrufen der Updates] " + stderr)
        return [], stderr
    try:
        updates = json.loads(stdout)
        if isinstance(updates, dict):  # Einzelnes Update als Dict
            updates = [updates]
        return updates, ""
    except:
        return [], "[Fehler beim Parsen der Updates]"

def install_selected_updates(kb_list):
    kb_str = ",".join(kb_list)
    command = (
        f"Import-Module PSWindowsUpdate; Get-WindowsUpdate | Where-Object {{$_.KB -in '{kb_str}'.Split(',')}} | Install-WindowsUpdate -AcceptAll -AutoReboot -Confirm:$false"
    )
    return run_powershell(command)

def fetch_hotfixes():
    command = "Get-HotFix | Select-Object -Property Description, HotFixID, InstalledOn | Format-List"
    return run_powershell(command)

def update_defender():
    command = "Update-MpSignature"
    return run_powershell(command)

class UpdateCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows Update Checker")
        self.root.geometry("900x650")

        if not has_internet():
            messagebox.showwarning("Keine Verbindung", "Es wurde keine Internetverbindung erkannt.")
            return

        if not is_admin():
            if not messagebox.askyesno("Administratorrechte benötigt", "Das Programm benötigt Administratorrechte, um Updates zu suchen. Jetzt neu starten mit Adminrechten?"):
                sys.exit()
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit()

        self.updates = []
        self.selected_kbs = []

        ttk.Label(root, text="Windows Update Checker", font=("Segoe UI", 16, "bold")).pack(pady=10)

        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=5)

        ttk.Button(btn_frame, text="🔍 Updates prüfen", command=self.check_updates).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="⬇️ Ausgewählte Updates installieren", command=self.install_selected).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="📋 Installierte Updates anzeigen", command=self.show_hotfixes).grid(row=0, column=2, padx=5)
        ttk.Button(btn_frame, text="🛡️ Defender updaten", command=self.update_defender).grid(row=0, column=3, padx=5)

        self.update_list = tk.Listbox(root, selectmode=tk.MULTIPLE, font=("Consolas", 10), height=10)
        self.update_list.pack(fill=tk.BOTH, expand=False, padx=10, pady=10)

        self.output_text = ScrolledText(root, wrap=tk.WORD, font=("Consolas", 10), height=15, state="disabled")
        self.output_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def log_output(self, text, color="black"):
        self.output_text.config(state="normal")
        self.output_text.insert(tk.END, text + "\n", color)
        self.output_text.tag_config("green", foreground="green")
        self.output_text.tag_config("yellow", foreground="orange")
        self.output_text.tag_config("red", foreground="red")
        self.output_text.config(state="disabled")
        self.output_text.see(tk.END)

    def check_updates(self):
        self.update_list.delete(0, tk.END)
        self.output_text.config(state="normal")
        self.output_text.delete("1.0", tk.END)
        self.output_text.config(state="disabled")
        self.log_output("🕐 Updates werden geprüft...", "yellow")

        def thread_func():
            updates, error = fetch_updates()
            if error:
                self.log_output(error, "red")
                return
            self.updates = updates
            if not updates:
                self.log_output("✅ Keine Updates gefunden.", "green")
            for up in updates:
                title = up.get("Title", "<Kein Titel>")
                kb = up.get("KB", "<KB?>")
                desc = up.get("Description", "")
                self.update_list.insert(tk.END, f"{kb} - {title} : {desc}")

        threading.Thread(target=thread_func, daemon=True).start()

    def install_selected(self):
        selected_indices = self.update_list.curselection()
        if not selected_indices:
            messagebox.showinfo("Keine Auswahl", "Bitte wähle mindestens ein Update aus.")
            return
        selected_kbs = [self.updates[i]["KB"] for i in selected_indices]
        self.log_output("⬇️ Installation gestartet für: " + ", ".join(selected_kbs), "yellow")

        def thread_func():
            stdout, stderr = install_selected_updates(selected_kbs)
            if stderr:
                self.log_output(stderr, "red")
            else:
                self.log_output(stdout, "green")

        threading.Thread(target=thread_func, daemon=True).start()

    def show_hotfixes(self):
        self.log_output("📋 Installierte Updates (HotFixes):", "yellow")
        def thread_func():
            stdout, stderr = fetch_hotfixes()
            if stderr:
                self.log_output(stderr, "red")
            else:
                self.log_output(stdout, "green")
        threading.Thread(target=thread_func, daemon=True).start()

    def update_defender(self):
        self.log_output("🛡️ Windows Defender wird aktualisiert...", "yellow")
        def thread_func():
            stdout, stderr = update_defender()
            if stderr:
                self.log_output(stderr, "red")
            else:
                self.log_output(stdout, "green")
        threading.Thread(target=thread_func, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = UpdateCheckerApp(root)
    root.mainloop()
