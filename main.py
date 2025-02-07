import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
from datetime import datetime, timedelta
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

class DateTimeSelector(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        # Date entry (YYYY-MM-DD)
        self.date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        self.date_entry = ttk.Entry(self, textvariable=self.date_var, width=12)
        self.date_entry.grid(row=0, column=0, padx=2)
        
        # Time entry (HH:MM:SS)
        self.time_var = tk.StringVar(value=datetime.now().strftime('%H:%M:%S'))
        self.time_entry = ttk.Entry(self, textvariable=self.time_var, width=10)
        self.time_entry.grid(row=0, column=1, padx=2)

    def get_datetime(self):
        try:
            return datetime.strptime(f"{self.date_var.get()} {self.time_var.get()}", 
                                   '%Y-%m-%d %H:%M:%S')
        except ValueError:
            return None

    def set_datetime(self, dt):
        self.date_var.set(dt.strftime('%Y-%m-%d'))
        self.time_var.set(dt.strftime('%H:%M:%S'))

class SysmonLogGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Sysmon Log Retriever (Admin)")
        self.root.geometry("700x500")
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Add admin indicator
        admin_label = ttk.Label(main_frame, text="Running as Administrator", foreground="green")
        admin_label.grid(row=0, column=0, columnspan=4, sticky=tk.W, pady=5)
        
        # Time selection frame
        time_frame = ttk.LabelFrame(main_frame, text="Time Range", padding="5")
        time_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
        
        # Start time
        ttk.Label(time_frame, text="Start Time:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        self.start_time = DateTimeSelector(time_frame)
        self.start_time.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        
        # End time
        ttk.Label(time_frame, text="End Time:").grid(row=0, column=2, sticky=tk.W, pady=5, padx=5)
        self.end_time = DateTimeSelector(time_frame)
        self.end_time.grid(row=0, column=3, sticky=tk.W, pady=5, padx=5)

        # Format hint
        hint_text = "Format: YYYY-MM-DD HH:MM:SS"
        ttk.Label(time_frame, text=hint_text, font=('', 8)).grid(
            row=1, column=0, columnspan=4, sticky=tk.W, padx=5)
        
        # Quick time buttons
        quick_frame = ttk.Frame(time_frame)
        quick_frame.grid(row=2, column=0, columnspan=4, pady=5)
        
        quick_times = [
            ("Last Hour", 3600),
            ("Last 6 Hours", 21600),
            ("Last 12 Hours", 43200),
            ("Last 24 Hours", 86400),
            ("Last 7 Days", 604800)
        ]
        
        for i, (text, seconds) in enumerate(quick_times):
            ttk.Button(quick_frame, text=text, 
                      command=lambda s=seconds: self.set_quick_time(s)).grid(
                          row=0, column=i, padx=5)
        
        # Output path selection
        path_frame = ttk.LabelFrame(main_frame, text="Output Settings", padding="5")
        path_frame.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(path_frame, text="Output Path:").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        self.output_path = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        path_entry = ttk.Entry(path_frame, textvariable=self.output_path, width=50)
        path_entry.grid(row=0, column=1, columnspan=2, sticky=tk.W, pady=5)
        
        browse_btn = ttk.Button(path_frame, text="Browse", command=self.browse_path)
        browse_btn.grid(row=0, column=3, sticky=tk.W, pady=5, padx=5)
        
        # PowerShell execution policy option
        self.bypass_policy = tk.BooleanVar(value=True)
        policy_check = ttk.Checkbutton(path_frame, text="Bypass ExecutionPolicy", 
                                     variable=self.bypass_policy)
        policy_check.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # Fetch button
        fetch_btn = ttk.Button(main_frame, text="Fetch Logs", command=self.fetch_logs)
        fetch_btn.grid(row=3, column=0, columnspan=4, pady=20)
        
        # Status label
        self.status_var = tk.StringVar()
        status_label = ttk.Label(main_frame, textvariable=self.status_var, wraplength=600)
        status_label.grid(row=4, column=0, columnspan=4, sticky=tk.W, pady=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, length=600, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)

    def set_quick_time(self, seconds):
        end = datetime.now()
        start = end - timedelta(seconds=seconds)
        self.start_time.set_datetime(start)
        self.end_time.set_datetime(end)

    def browse_path(self):
        directory = filedialog.askdirectory(initialdir=self.output_path.get())
        if directory:
            self.output_path.set(directory)

    def validate_times(self):
        start = self.start_time.get_datetime()
        end = self.end_time.get_datetime()
        
        if not start or not end:
            messagebox.showerror("Error", "Invalid date/time format.\nUse: YYYY-MM-DD HH:MM:SS")
            return False
            
        if start > end:
            messagebox.showerror("Error", "Start time must be before end time")
            return False
            
        return True

    def fetch_logs(self):
        if not self.validate_times():
            return
            
        start_time = self.start_time.get_datetime().strftime('%Y-%m-%dT%H:%M:%S')
        end_time = self.end_time.get_datetime().strftime('%Y-%m-%dT%H:%M:%S')
        
        output_file = os.path.join(
            self.output_path.get(), 
            f"sysmon_logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )

        # Construct PowerShell command
        ps_command = (
            f'Get-WinEvent -LogName "Microsoft-Windows-Sysmon/Operational" '
            f'-FilterHashTable @{{LogName="Microsoft-Windows-Sysmon/Operational";'
            f'StartTime="{start_time}";'
            f'EndTime="{end_time}"}} | Export-Csv -Path "{output_file}" -NoTypeInformation'
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