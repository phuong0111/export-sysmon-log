import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
from datetime import datetime
import os
import sys
import ctypes

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

class SysmonLogGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Sysmon Log Retriever (Admin)")
        self.root.geometry("600x400")
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Add admin indicator
        admin_label = ttk.Label(main_frame, text="Running as Administrator", foreground="green")
        admin_label.grid(row=0, column=0, columnspan=4, sticky=tk.W, pady=5)
        
        # Time range selection
        ttk.Label(main_frame, text="Time Range:").grid(row=1, column=0, sticky=tk.W, pady=5)
        
        self.time_value = tk.StringVar(value="3600")
        self.time_entry = ttk.Entry(main_frame, textvariable=self.time_value, width=10)
        self.time_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        self.time_unit = tk.StringVar(value="seconds")
        time_unit_combo = ttk.Combobox(main_frame, textvariable=self.time_unit, 
                                     values=["seconds", "minutes", "hours"], 
                                     width=8, state="readonly")
        time_unit_combo.grid(row=1, column=2, sticky=tk.W, pady=5, padx=5)
        
        # Output path selection
        ttk.Label(main_frame, text="Output Path:").grid(row=2, column=0, sticky=tk.W, pady=5)
        
        self.output_path = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        path_entry = ttk.Entry(main_frame, textvariable=self.output_path, width=50)
        path_entry.grid(row=2, column=1, columnspan=2, sticky=tk.W, pady=5)
        
        browse_btn = ttk.Button(main_frame, text="Browse", command=self.browse_path)
        browse_btn.grid(row=2, column=3, sticky=tk.W, pady=5, padx=5)
        
        # PowerShell execution policy option
        self.bypass_policy = tk.BooleanVar(value=True)
        policy_check = ttk.Checkbutton(main_frame, text="Bypass ExecutionPolicy", 
                                     variable=self.bypass_policy)
        policy_check.grid(row=3, column=1, columnspan=2, sticky=tk.W, pady=5)
        
        # Fetch button
        fetch_btn = ttk.Button(main_frame, text="Fetch Logs", command=self.fetch_logs)
        fetch_btn.grid(row=4, column=1, sticky=tk.W, pady=20)
        
        # Status label
        self.status_var = tk.StringVar()
        status_label = ttk.Label(main_frame, textvariable=self.status_var, wraplength=500)
        status_label.grid(row=5, column=0, columnspan=4, sticky=tk.W, pady=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, length=400, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)

    def browse_path(self):
        directory = filedialog.askdirectory(initialdir=self.output_path.get())
        if directory:
            self.output_path.set(directory)

    def convert_to_seconds(self):
        try:
            value = float(self.time_value.get())
            unit = self.time_unit.get()
            
            if unit == "minutes":
                return value * 60
            elif unit == "hours":
                return value * 3600
            return value
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number for time range")
            return None

    def fetch_logs(self):
        seconds = self.convert_to_seconds()
        if seconds is None:
            return

        output_file = os.path.join(
            self.output_path.get(), 
            f"sysmon_logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )

        # Construct PowerShell command
        ps_command = (
            f'Get-WinEvent -LogName "Microsoft-Windows-Sysmon/Operational" '
            f'-FilterHashTable @{{ LogName = "Microsoft-Windows-Sysmon/Operational";'
            f'StartTime = (Get-Date).AddSeconds(-{seconds});'
            f'EndTime = (Get-Date) }} | Export-Csv -Path "{output_file}" -NoTypeInformation'
        )

        self.status_var.set("Fetching logs...")
        self.progress.start()
        self.root.update()

        try:
            # Prepare PowerShell command with execution policy bypass if selected
            powershell_args = ["powershell"]
            if self.bypass_policy.get():
                powershell_args.extend(["-ExecutionPolicy", "Bypass"])
            powershell_args.extend(["-Command", ps_command])

            # Execute PowerShell command
            process = subprocess.run(
                powershell_args,
                capture_output=True,
                text=True,
                check=True
            )
            
            self.status_var.set(f"Logs successfully exported to:\n{output_file}")
            messagebox.showinfo("Success", "Logs have been successfully exported!")
            
        except subprocess.CalledProcessError as e:
            error_msg = f"Error fetching logs: {e.stderr}"
            self.status_var.set(error_msg)
            messagebox.showerror("Error", error_msg)
            
        finally:
            self.progress.stop()

def main():
    if not is_admin():
        messagebox.showerror("Admin Required", "This application needs administrator privileges to access event logs.")
        run_as_admin()
        sys.exit()

    root = tk.Tk()
    app = SysmonLogGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()