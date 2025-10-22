import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta
import threading
import subprocess
import sys
import os

# Fix DPI scaling issues on Windows
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

class DateRangeGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Weekly Mezz Automated")
        self.root.geometry("500x350")
        self.root.resizable(False, False)
        
        # Configure style for better appearance
        self.setup_styles()
        
        # Center the window
        self.center_window()
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="30")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="주간발행내역 자동화", 
                               font=("Segoe UI", 18, "bold"), style="Title.TLabel")
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 30))
        
        # From Date
        ttk.Label(main_frame, text="시작일 (YYYYMMDD):").grid(row=1, column=0, sticky=tk.W, pady=8)
        self.from_date_var = tk.StringVar(value="20251015")
        from_date_entry = ttk.Entry(main_frame, textvariable=self.from_date_var, width=20, font=("Segoe UI", 10))
        from_date_entry.grid(row=1, column=1, sticky=tk.W, padx=(15, 0), pady=8)
        
        # To Date
        ttk.Label(main_frame, text="종료일 (YYYYMMDD):").grid(row=2, column=0, sticky=tk.W, pady=8)
        self.to_date_var = tk.StringVar(value="20251021")
        to_date_entry = ttk.Entry(main_frame, textvariable=self.to_date_var, width=20, font=("Segoe UI", 10))
        to_date_entry.grid(row=2, column=1, sticky=tk.W, padx=(15, 0), pady=8)
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=25)
        
        # Run button
        self.run_button = ttk.Button(button_frame, text="실행", command=self.run_script, 
                                   style="Accent.TButton", width=10)
        self.run_button.pack(side=tk.LEFT, padx=(0, 15))
        
        # Exit button
        exit_button = ttk.Button(button_frame, text="종료", command=self.root.quit, 
                               style="TButton", width=10)
        exit_button.pack(side=tk.LEFT)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate', style="TProgressbar")
        self.progress.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=15)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="실행 준비 완료", foreground="#0d7efb", 
                                    font=("Segoe UI", 9))
        self.status_label.grid(row=5, column=0, columnspan=2, pady=8)
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
    
    def setup_styles(self):
        """Configure custom styles for better appearance"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Define custom colors
        blue = "#0d7efb"      # (13, 126, 251)
        grey = "#616365"      # (97, 99, 101)
        white = "#f5f5f5"     # (245, 245, 245)
        
        # Configure title style
        style.configure("Title.TLabel", 
                       font=("Segoe UI", 18, "bold"),
                       foreground=blue)
        
        # Configure label style
        style.configure("TLabel", 
                       font=("Segoe UI", 10),
                       foreground=grey,
                       background=white)
        
        # Configure entry style
        style.configure("TEntry", 
                       font=("Segoe UI", 10),
                       fieldbackground="white",
                       borderwidth=2,
                       relief="solid",
                       bordercolor=grey)
        
        # Configure button styles
        style.configure("Accent.TButton",
                       font=("Segoe UI", 11, "bold"),
                       foreground="white",
                       background=blue,
                       borderwidth=0,
                       focuscolor="none")
        
        style.map("Accent.TButton",
                 background=[("active", "#0a6bdf"),
                           ("pressed", "#0859c7")])
        
        # Configure regular button
        style.configure("TButton",
                       font=("Segoe UI", 10, "bold"),
                       foreground=grey,
                       background=white,
                       borderwidth=1,
                       bordercolor=grey,
                       focuscolor="none")
        
        style.map("TButton",
                 background=[("active", "#e8e8e8"),
                           ("pressed", "#dcdcdc")])
        
        # Configure progress bar
        style.configure("TProgressbar",
                       background=blue,
                       troughcolor=white,
                       borderwidth=0,
                       lightcolor=blue,
                       darkcolor=blue)
        
        # Configure frame background
        style.configure("TFrame",
                       background=white)
    
    def center_window(self):
        """Center the window on the screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def validate_date(self, date_string):
        """Validate date format YYYYMMDD"""
        try:
            datetime.strptime(date_string, "%Y%m%d")
            return True
        except ValueError:
            return False
    
    def run_script(self):
        """Run the main script with the provided dates"""
        from_date = self.from_date_var.get().strip()
        to_date = self.to_date_var.get().strip()
        
        # Validate dates
        if not self.validate_date(from_date):
            messagebox.showerror("Error", "Invalid from date format. Use YYYYMMDD")
            return
        
        if not self.validate_date(to_date):
            messagebox.showerror("Error", "Invalid to date format. Use YYYYMMDD")
            return
        
        # Check if from_date is before to_date
        if datetime.strptime(from_date, "%Y%m%d") > datetime.strptime(to_date, "%Y%m%d"):
            messagebox.showerror("Error", "From date must be before or equal to to date")
            return
        
        # Disable run button and start progress
        self.run_button.config(state='disabled')
        self.progress.start()
        self.status_label.config(text="실행 중...", foreground="#0d7efb")
        
        # Run script in separate thread to prevent GUI freezing
        thread = threading.Thread(target=self.execute_script, args=(from_date, to_date))
        thread.daemon = True
        thread.start()
    
    def execute_script(self, from_date, to_date):
        """Execute the data processing directly"""
        try:
            # Import required modules
            import sys
            import os
            
            # Import and run the data processing functions directly
            import main
            
            # Run the data processing
            reports = main.get_weekly_reports(from_date, to_date)
            
            filter_words = [
                '감자', '증자', '합병', '분할', '해산', '증여',
                '자기', '자본', '자산', '담보', 
                '양수도', '양수', '양도', '처분', 
                '선택권', '소조',  '보증', 
            ]

            filtered_reports = []
            for report in reports:
                if any(word in report['report_nm'] for word in filter_words): continue
                contain_keys = ['stock_code', 'report_nm', 'corp_code', 'corp_name', 'rcept_no', 'corp_cls']
                filtered_report = {key: report[key] for key in contain_keys if key in report}
                filtered_reports.append(filtered_report)
            
            # Run the table processing
            main.table_to_xlsx(filtered_reports)
            
            # Create a mock result object
            class MockResult:
                def __init__(self):
                    self.returncode = 0
                    self.stderr = ""
            
            result = MockResult()
            self.root.after(0, self.script_finished, result)
            
        except Exception as e:
            self.root.after(0, self.script_error, str(e))
    
    def script_finished(self, result):
        """Handle script completion"""
        self.progress.stop()
        self.run_button.config(state='normal')
        
        if result.returncode == 0:
            self.status_label.config(text="실행 완료!", foreground="#0d7efb")
            messagebox.showinfo("완료", "스크립트가 성공적으로 완료되었습니다!\noutput.xlsx 파일을 확인하세요.")
        else:
            self.status_label.config(text="오류 발생", foreground="#616365")
            messagebox.showerror("오류", f"스크립트 실행 중 오류가 발생했습니다:\n{result.stderr}")
    
    def script_error(self, error_msg):
        """Handle script error"""
        self.progress.stop()
        self.run_button.config(state='normal')
        self.status_label.config(text="오류 발생", foreground="#616365")
        messagebox.showerror("오류", f"스크립트 실행 실패:\n{error_msg}")

def main():
    root = tk.Tk()
    
    # Configure style
    style = ttk.Style()
    style.theme_use('clam')
    
    # Create and run the GUI
    app = DateRangeGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
