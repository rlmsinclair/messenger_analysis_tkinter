import tkinter as tk
from tkinter import ttk, messagebox

import anthropic
import customtkinter as ctk
import json
import time
import threading
import queue
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC, wait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from datetime import datetime
import sys
from pathlib import Path

# Set up CustomTkinter appearance
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# Only class
class ModernMessengerExporter:
    def __init__(self):
        # Existing initialization code remains the same
        self.root = ctk.CTk()
        self.root.title("Facebook Messenger Analyzer")
        self.root.geometry("800x800")
        self.root.minsize(800, 800)

        # Add API key variable
        self.api_key = ctk.StringVar()

        # Multiple thread control flags
        self.selenium_running = threading.Event()
        self.export_running = threading.Event()

        # Queues for thread communication
        self.message_queue = queue.Queue()
        self.command_queue = queue.Queue()

        # Thread handles
        self.selenium_thread = None
        self.export_thread = None
        self.cleanup_thread = None
        self.analysis_thread = None

        # Selenium driver
        self.driver = None
        self.driver_lock = threading.Lock()

        self.setup_variables()
        self.create_gui()
        self.process_queues()
        self.processed_messages = set()

    def setup_variables(self):
        """Initialize all variables needed for the application"""
        self.current_step = 1
        self.login_verified = False
        self.processed_messages = set()
        self.last_scroll_time = 0
        self.same_message_count = 0

        # GUI state variables
        self.login_method = ctk.StringVar(value="manual")
        self.chat_type = ctk.StringVar(value="individual")
        self.output_path = ctk.StringVar(value=str(Path.home() / "Downloads" / "conversation.txt"))

    def process_queues(self):
        """Process message and command queues"""
        try:
            # Process status messages
            while True:
                try:
                    msg = self.message_queue.get_nowait()
                    if msg.get('type') == 'status':
                        self._update_status(msg['message'], msg['level'])
                    elif msg.get('type') == 'complete':
                        self._handle_completion()
                    self.message_queue.task_done()
                except queue.Empty:
                    break

            # Process commands
            while True:
                try:
                    cmd = self.command_queue.get_nowait()
                    if cmd.get('type') == 'enable_button':
                        self.export_button.configure(state="normal")
                    elif cmd.get('type') == 'update_button':
                        self.export_button.configure(**cmd['properties'])
                    self.command_queue.task_done()
                except queue.Empty:
                    break
        finally:
            self.root.after(100, self.process_queues)

    def create_gui(self):
        """Updated GUI creation to include step 4"""
        # Create main container
        self.main_container = ctk.CTkFrame(self.root)
        self.main_container.pack(fill="both", expand=True, padx=15, pady=15)

        # Header
        self.create_header()

        # Progress Steps
        self.create_progress_steps()

        # Content Area
        self.content_area = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.content_area.pack(fill="both", expand=True, pady=10)

        # Create all step frames
        self.step_frames = {
            1: self.create_step1_frame(),
            2: self.create_step2_frame(),
            3: self.create_step3_frame(),
            4: self.create_step4_frame()
        }

        # Navigation Buttons
        self.create_navigation()

        # Show first step
        self.show_step(1)

    def create_header(self):
        """Create the header section"""
        header_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 20))

        title = ctk.CTkLabel(
            header_frame,
            text="Facebook Messenger Analyzer",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title.pack()

        subtitle = ctk.CTkLabel(
            header_frame,
            text="Gain insight about yourself and your friends",
            text_color="gray"
        )
        subtitle.pack()

    def create_progress_steps(self):
        """Updated progress steps to include analysis step"""
        self.progress_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.progress_frame.pack(fill="x", pady=(0, 20))

        # Configure grid columns
        for i in range(7):  # Increased for 4 steps
            self.progress_frame.grid_columnconfigure(i, weight=1 if i % 2 else 0)

        # Step circles and labels
        self.step_indicators = {}
        steps = [("1", "Login"), ("2", "Chat Type"), ("3", "Export"), ("4", "Analysis")]

        for i, (num, label) in enumerate(steps):
            # Create circle
            circle = ctk.CTkButton(
                self.progress_frame,
                text=num,
                width=30,
                height=30,
                corner_radius=15,
                state="disabled"
            )
            circle.grid(row=0, column=i * 2)

            # Create label
            step_label = ctk.CTkLabel(self.progress_frame, text=label)
            step_label.grid(row=1, column=i * 2)

            self.step_indicators[int(num)] = (circle, step_label)

            # Add connector line except for last step
            if i < len(steps) - 1:
                line = ctk.CTkFrame(
                    self.progress_frame,
                    height=2,
                    fg_color="gray75"
                )
                line.grid(row=0, column=i * 2 + 1, sticky="ew", padx=10)

    def start_analysis(self):
        """Begin the chat analysis process"""
        if not os.path.exists(self.output_path.get()):
            messagebox.showerror("Error", "Chat export file not found")
            return

        self.analyze_button.configure(state="disabled")
        self.analysis_thread = threading.Thread(target=self._perform_analysis)
        self.analysis_thread.daemon = True
        self.analysis_thread.start()

    def _perform_analysis(self):
        """Perform the chat analysis using Flask API"""
        try:
            # Read the chat file
            with open(self.output_path.get(), 'r', encoding='utf-8') as f:
                chat_content = f.read()

            # Update status
            self._update_analysis_status("Analyzing chat content\n")

            # Make request to Flask API
            import requests
            response = requests.post(
                'https://messenger-analysis-api-k63dd.ondigitalocean.app/analyze',
                json={'chat_content': chat_content},
                headers={'Content-Type': 'application/json'}
            )

            if response.status_code == 200:
                # Get the analysis from response
                analysis = response.json()['analysis']

                # Update GUI with analysis
                self._update_analysis_status(analysis)
                self._update_analysis_status("\nAnalysis complete!")
            else:
                error_message = response.json().get('error', 'Unknown error occurred')
                self._update_analysis_status(f"\nError during analysis: {error_message}")

        except requests.exceptions.ConnectionError:
            self._update_analysis_status(
                "\nError: Could not connect to analysis server. Please make sure the server is running.")
        except Exception as e:
            self._update_analysis_status(f"\nError during analysis: {str(e)}")
        finally:
            self.root.after(0, lambda: self.analyze_button.configure(state="normal"))

    def _update_analysis_status(self, message):
        """Update the analysis status text in a thread-safe way"""

        def update():
            self.analysis_text.configure(state="normal")
            self.analysis_text.insert("end", message + "\n")
            self.analysis_text.configure(state="disabled")
            self.analysis_text.see("end")

        self.root.after(0, update)

    def create_step1_frame(self):
        """Create the login method selection frame"""
        frame = ctk.CTkFrame(self.content_area, fg_color="transparent")

        # Title
        title = ctk.CTkLabel(
            frame,
            text="How would you like to login to Facebook?",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        title.pack(pady=(0, 20))

        # Login Method Options
        methods_frame = ctk.CTkFrame(frame, fg_color="transparent")
        methods_frame.pack(fill="x")

        # Manual Login Option
        manual_frame = ctk.CTkFrame(methods_frame)
        manual_frame.pack(fill="x", pady=5)

        manual_radio = ctk.CTkRadioButton(
            manual_frame,
            text="Login with Facebook",
            variable=self.login_method,
            value="manual"
        )
        manual_radio.pack(pady=10, padx=10)

        manual_description = ctk.CTkLabel(
            manual_frame,
            text="Recommended: Simply login to Facebook in the browser window",
            text_color="gray"
        )
        manual_description.pack(pady=(0, 10), padx=10)

        # Cookies Login Option
        cookies_frame = ctk.CTkFrame(methods_frame)
        cookies_frame.pack(fill="x", pady=5)

        cookies_radio = ctk.CTkRadioButton(
            cookies_frame,
            text="Use Browser Cookies",
            variable=self.login_method,
            value="cookies"
        )
        cookies_radio.pack(pady=10, padx=10)

        cookies_description = ctk.CTkLabel(
            cookies_frame,
            text="Advanced: Paste your Facebook cookies here",
            text_color="gray"
        )
        cookies_description.pack(pady=(0, 10), padx=10)

        self.cookies_textbox = ctk.CTkTextbox(
            cookies_frame,
            height=100,
            state="disabled"
        )

        # Enable/disable cookies textbox based on selection
        def toggle_cookies(*args):
            if self.login_method.get() == "cookies":
                self.cookies_textbox.pack(fill="x", padx=10, pady=(0, 10))
                self.cookies_textbox.configure(state="normal")
            else:
                self.cookies_textbox.pack_forget()
                self.cookies_textbox.configure(state="disabled")

        self.login_method.trace("w", toggle_cookies)

        return frame

    def create_step2_frame(self):
        """Create the chat type selection frame"""
        frame = ctk.CTkFrame(self.content_area, fg_color="transparent")

        # Title
        title = ctk.CTkLabel(
            frame,
            text="What type of chat do you want to export?",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        title.pack(pady=(0, 20))

        # Chat Type Options
        types_frame = ctk.CTkFrame(frame, fg_color="transparent")
        types_frame.pack(fill="x")

        # Individual Chat Option
        individual_frame = ctk.CTkFrame(types_frame)
        individual_frame.pack(fill="x", pady=5)

        individual_radio = ctk.CTkRadioButton(
            individual_frame,
            text="Individual Chat",
            variable=self.chat_type,
            value="individual"
        )
        individual_radio.pack(pady=10, padx=10)

        individual_description = ctk.CTkLabel(
            individual_frame,
            text="A conversation between you and one other person",
            text_color="gray"
        )
        individual_description.pack(pady=(0, 10), padx=10)

        # Group Chat Option
        group_frame = ctk.CTkFrame(types_frame)
        group_frame.pack(fill="x", pady=5)

        group_radio = ctk.CTkRadioButton(
            group_frame,
            text="Group Chat",
            variable=self.chat_type,
            value="group"
        )
        group_radio.pack(pady=10, padx=10)

        group_description = ctk.CTkLabel(
            group_frame,
            text="A conversation with multiple people",
            text_color="gray"
        )
        group_description.pack(pady=(0, 10), padx=10)

        return frame

    def create_step3_frame(self):
        """Create the export frame"""
        frame = ctk.CTkFrame(self.content_area, fg_color="transparent")

        # Title
        title = ctk.CTkLabel(
            frame,
            text="Ready to Export",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        title.pack(pady=(0, 10))

        # Instructions
        instructions_frame = ctk.CTkFrame(frame)
        instructions_frame.pack(fill="x", pady=5)

        instructions_title = ctk.CTkLabel(
            instructions_frame,
            text="Instructions",
            font=ctk.CTkFont(weight="bold")
        )
        instructions_title.pack(pady=(5, 5))

        instructions = [
            "1. Click 'Start Export' below",
            "2. Select the conversation you want to save in the Facebook window",
            "3. (Optional for better performance) Zoom out as far as possible using 'Ctrl -' or 'Command -'",
            "4. Click confirm. Wait while we save your messages. You may be prompted to scroll up in the chat.",
            "5. Your file will be saved automatically."
        ]

        for instruction in instructions:
            instruction_label = ctk.CTkLabel(
                instructions_frame,
                text=instruction
            )
            instruction_label.pack(pady=(0, 2))

        # Status Display
        status_frame = ctk.CTkFrame(frame)
        status_frame.pack(fill="both", expand=True, pady=5)

        self.status_text = ctk.CTkTextbox(
            status_frame,
            height=150,
            state="disabled"
        )
        self.status_text.pack(fill="both", expand=True, padx=10, pady=5)

        return frame

    def create_step4_frame(self):
        """Create the analysis frame without API key input"""
        frame = ctk.CTkFrame(self.content_area, fg_color="transparent")

        # Title
        title = ctk.CTkLabel(
            frame,
            text="Analyze Chat",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        title.pack(pady=(0, 10))

        # Analysis Status
        self.analysis_text = ctk.CTkTextbox(
            frame,
            height=400,
            state="disabled"
        )
        self.analysis_text.pack(fill="both", expand=True, pady=10)

        # Analyze Button
        self.analyze_button = ctk.CTkButton(
            frame,
            text="Analyze Chat",
            command=self.start_analysis
        )
        self.analyze_button.pack(pady=10)

        return frame

    def create_navigation(self):
        """Create navigation buttons"""
        nav_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        nav_frame.pack(fill="x", pady=(10, 0))

        self.back_button = ctk.CTkButton(
            nav_frame,
            text="Back",
            command=self.go_back,
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90")
        )
        self.back_button.pack(side="left")

        # Create export button
        self.export_button = ctk.CTkButton(
            nav_frame,
            text="Start Export",
            command=self.toggle_export,
            fg_color=("blue", "blue"),
            hover_color=("dark blue", "dark blue")
        )

        # Create continue button
        self.next_button = ctk.CTkButton(
            nav_frame,
            text="Continue",
            command=self.go_next
        )

        # Initially show next button
        self.next_button.pack(side="right")

        # Initially disable back button
        self.back_button.configure(state="disabled")

    def show_step(self, step_number):
        """Updated show_step method to include step 4"""
        # Hide all frames
        for frame in self.step_frames.values():
            frame.pack_forget()

        # Show requested frame
        self.step_frames[step_number].pack(fill="both", expand=True)

        # Update progress indicators
        for step, (circle, label) in self.step_indicators.items():
            if step < step_number:
                circle.configure(fg_color="green", hover_color="green")
            elif step == step_number:
                circle.configure(fg_color="blue", hover_color="blue")
            else:
                circle.configure(fg_color="gray75", hover_color="gray75")

        # Update navigation buttons
        self.back_button.configure(state="normal" if step_number > 1 else "disabled")

        # Handle button visibility
        self.next_button.pack_forget()
        self.export_button.pack_forget()
        self.analyze_button.pack_forget()

        if step_number == 3:
            self.export_button.pack(side="right")
        elif step_number == 4:
            self.analyze_button.pack(side="right")
        else:
            self.next_button.pack(side="right")

        self.current_step = step_number

    def go_back(self):
        """Navigate to previous step"""
        if self.current_step > 1:
            self.show_step(self.current_step - 1)

    def go_next(self):
        """Navigate to next step"""
        if self.current_step < 3:
            self.show_step(self.current_step + 1)

    def _update_status(self, message, level="info"):
        """Update status text in a thread-safe way"""
        self.status_text.configure(state="normal")
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_text.insert("end", f"[{timestamp}] {message}\n")
        self.status_text.configure(state="disabled")
        self.status_text.see("end")

    def _handle_completion(self):
        """Handle export completion and move to analysis step"""
        self.selenium_running.clear()
        self.export_running.clear()
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None
        self._reset_export_button()
        # Move to analysis step
        self.root.after(0, lambda: self.show_step(4))

    def _reset_export_button(self):
        """Reset export button in a thread-safe way"""
        self.export_button.configure(
            text="Start Export",
            state="normal",
            fg_color=("blue", "blue"),
            hover_color=("dark blue", "dark blue")
        )

    def toggle_export(self):
        """Toggle between starting and stopping export"""
        if not self.selenium_running.is_set() and not self.export_running.is_set():
            self.start_export()
        else:
            self.stop_export()

    def start_export(self):
        """Start the export process in separate threads"""
        if self.selenium_thread and self.selenium_thread.is_alive():
            return

        # Set control flags
        self.selenium_running.set()
        self.export_running.set()

        # Update button state
        self.export_button.configure(
            text="Stop Export",
            fg_color="red",
            hover_color="dark red"
        )

        # Start Selenium thread
        self.selenium_thread = threading.Thread(target=self.initialize_selenium)
        self.selenium_thread.daemon = True
        self.selenium_thread.start()

    def stop_export(self):
        """Gracefully stop all export processes and move to analysis"""
        # Move to step 4 immediately
        self.show_step(4)

        # Clear control flags
        self.selenium_running.clear()
        self.export_running.clear()

        # Disable button during cleanup
        self.export_button.configure(state="disabled")

        # Start cleanup in separate thread
        self.cleanup_thread = threading.Thread(target=self._cleanup)
        self.cleanup_thread.daemon = True
        self.cleanup_thread.start()

        self.message_queue.put({
            'type': 'status',
            'message': 'Cleaning up resources in background...',
            'level': 'info'
        })

    def _cleanup(self):
        """Clean up resources in separate thread"""
        try:
            with self.driver_lock:
                if self.driver:
                    self.driver.quit()
                    self.driver = None
        except Exception as e:
            self.message_queue.put({
                'type': 'status',
                'message': f'Error during cleanup: {str(e)}',
                'level': 'error'
            })
        finally:
            self.command_queue.put({
                'type': 'update_button',
                'properties': {
                    'text': 'Start Export',
                    'state': 'normal',
                    'fg_color': ("blue", "blue"),
                    'hover_color': ("dark blue", "dark blue")
                }
            })

    def initialize_selenium(self):
        """Initialize Selenium in separate thread"""
        try:
            with self.driver_lock:
                self.driver = webdriver.Chrome()
                self.driver.maximize_window()
                self.driver.get("https://www.facebook.com")

            # Handle login
            if self.login_method.get() == "cookies":
                self._handle_cookie_login()
            else:
                self._handle_manual_login()

            if self.selenium_running.is_set():
                self.message_queue.put({
                    'type': 'status',
                    'message': 'Please select the conversation to export.',
                    'level': 'info'
                })
                self.root.after(1000, self.create_confirmation_popup)

        except Exception as e:
            self.message_queue.put({
                'type': 'status',
                'message': f'Error: {str(e)}',
                'level': 'error'
            })
            self._cleanup()

    def _handle_cookie_login(self):
        """Handle cookie-based login with proper error handling"""
        try:
            cookies_str = self.cookies_textbox.get("1.0", "end").strip()
            if not cookies_str:
                raise ValueError("No cookies provided")

            cookies = json.loads(cookies_str)
            if not isinstance(cookies, list):
                raise ValueError("Cookies must be a JSON array")

            for cookie in cookies:
                required_fields = ['name', 'value', 'domain']
                missing_fields = [field for field in required_fields if field not in cookie]
                if missing_fields:
                    raise ValueError(f"Cookie missing required fields: {', '.join(missing_fields)}")

                # Construct properly formatted cookie
                cookie_dict = {
                    'name': cookie['name'],
                    'value': cookie['value'],
                    'domain': cookie['domain'],
                    'path': cookie.get('path', '/'),
                }

                # Optional fields
                if 'expirationDate' in cookie:
                    cookie_dict['expiry'] = int(cookie['expirationDate'])
                if 'secure' in cookie:
                    cookie_dict['secure'] = cookie['secure']
                if 'httpOnly' in cookie:
                    cookie_dict['httpOnly'] = cookie['httpOnly']

                self.driver.add_cookie(cookie_dict)

            self.driver.refresh()

            # Verify login success
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "[aria-label='Facebook']"))
            )

            self.message_queue.put({
                'type': 'status',
                'message': 'Successfully logged in using cookies',
                'level': 'info'
            })

            self.driver.get("https://www.facebook.com/messages/t/")

        except json.JSONDecodeError as e:
            error_msg = f"Invalid cookie format: {str(e)}"
            self.message_queue.put({
                'type': 'status',
                'message': error_msg + ". Using manual login.",
                'level': 'warning'
            })
            self._handle_manual_login()

        except ValueError as e:
            error_msg = f"Cookie validation error: {str(e)}"
            self.message_queue.put({
                'type': 'status',
                'message': error_msg + ". Using manual login.",
                'level': 'warning'
            })
            self._handle_manual_login()

        except Exception as e:
            error_msg = f"Cookie login error: {str(e)}"
            self.message_queue.put({
                'type': 'status',
                'message': error_msg + ". Using manual login.",
                'level': 'warning'
            })
            self._handle_manual_login()

    def _handle_manual_login(self):
        """Handle manual login process"""
        self.message_queue.put({
            'type': 'status',
            'message': 'Please log in to Facebook in the browser window...',
            'level': 'info'
        })

        # Wait for login
        start_time = time.time()
        while time.time() - start_time < 300 and self.selenium_running.is_set():  # 5 minute timeout
            try:
                if len(self.driver.find_elements(By.CSS_SELECTOR, "[aria-label='Facebook']")) > 0:
                    break
            except:
                pass
            time.sleep(2)
        self.driver.get("https://www.facebook.com/messages/t/")

    def create_confirmation_popup(self):
        """Create confirmation dialog"""
        if not self.selenium_running.is_set():
            return

        popup = ctk.CTkToplevel(self.root)
        popup.title("Confirm Chat Selection")
        popup.geometry("600x300")

        # Center popup
        popup.update_idletasks()
        x = (popup.winfo_screenwidth() - popup.winfo_width()) // 2
        y = (popup.winfo_screenheight() - popup.winfo_height()) // 2
        popup.geometry(f"+{x}+{y}")

        popup.transient(self.root)
        popup.grab_set()

        ctk.CTkLabel(
            popup,
            text="Please select the chat you wish to export in the browser.\nPlease then zoom out as far as possible using 'Ctrl -' or 'Command -'",
            font=ctk.CTkFont(size=16),
            justify="center"
        ).pack(pady=(50, 50))

        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=20)

        ctk.CTkButton(
            btn_frame,
            text="Confirm",
            command=lambda: self._handle_confirmation(popup, True),  # Removed underscore
            fg_color="green",
            hover_color="dark green"
        ).pack(side="left", padx=10)

        ctk.CTkButton(
            btn_frame,
            text="Cancel",
            command=lambda: self._handle_confirmation(popup, False),  # Removed underscore
            fg_color="red",
            hover_color="dark red"
        ).pack(side="left", padx=10)

    def _handle_confirmation(self, popup, confirmed):
        """Handle confirmation response"""
        popup.destroy()
        if confirmed and self.selenium_running.is_set():
            self.begin_message_export()
        else:
            self.stop_export()

    def begin_message_export(self):
        """Start message export in separate thread"""
        if not self.export_running.is_set():
            return

        # Start export thread
        self.export_thread = threading.Thread(target=self._export_messages)
        self.export_thread.daemon = True
        self.export_thread.start()

    def get_message_identifier(self, message_element):
        """Generate a unique identifier for a message using its text content and class names"""
        try:
            # Get the full class name string
            class_names = message_element.get_attribute("class")
            # Get the text content - encode to handle emojis
            text_content = message_element.text.encode('utf-8', 'ignore').decode('utf-8').strip()
            # Get parent div's class names for additional context
            parent_classes = message_element.find_element(By.XPATH, "..").get_attribute("class")

            # Create a compound identifier
            identifier = f"{class_names}|{parent_classes}|{text_content[:50]}"  # Use first 50 chars of text
            return identifier
        except Exception as e:
            print(f"Debug: Error creating message identifier: {e}")
            return None

    def _get_message_color(self, element):
        try:
            color = self.driver.execute_script("""
                var element = arguments[0];
                var bgColor = window.getComputedStyle(element).backgroundColor;
                if (bgColor === 'rgba(0, 0, 0, 0)' || !bgColor) {
                    var parent = element.parentElement;
                    while (parent && window.getComputedStyle(parent).backgroundColor === 'rgba(0, 0, 0, 0)') {
                        parent = parent.parentElement;
                    }
                    return parent ? window.getComputedStyle(parent).backgroundColor : null;
                }
                return bgColor;
            """, element)
            return color
        except Exception as e:
            print(f"Debug: Error getting color: {e}")
            return None

    def create_scroll_warning_popup(self):
        """Create warning popup for scroll issues"""
        popup = ctk.CTkToplevel(self.root)
        popup.title("Scroll Warning")
        popup.geometry("400x200")

        # Center popup
        popup.update_idletasks()
        x = (popup.winfo_screenwidth() - popup.winfo_width()) // 2
        y = (popup.winfo_screenheight() - popup.winfo_height()) // 2
        popup.geometry(f"+{x}+{y}")

        popup.transient(self.root)
        popup.grab_set()

        # Warning icon and message
        warning_frame = ctk.CTkFrame(popup, fg_color="transparent")
        warning_frame.pack(expand=True, fill="both", padx=20, pady=20)

        message_label = ctk.CTkLabel(
            warning_frame,
            text="Unable to load more messages.\nPlease either:\n\n1. Scroll up manually in the chat window\n2. Stop the export",
            font=ctk.CTkFont(size=14),
            justify="center"
        )
        message_label.pack(pady=(20, 30))

        # Buttons
        button_frame = ctk.CTkFrame(warning_frame, fg_color="transparent")
        button_frame.pack(fill="x")

        ctk.CTkButton(
            button_frame,
            text="Stop Export",
            command=lambda: self._handle_scroll_warning(popup, True),
            fg_color="red",
            hover_color="dark red"
        ).pack(side="left", expand=True, padx=10)

        ctk.CTkButton(
            button_frame,
            text="Continue",
            command=lambda: self._handle_scroll_warning(popup, False),
        ).pack(side="left", expand=True, padx=10)

    def _handle_scroll_warning(self, popup, stop=False):
        """Handle the user's response to the scroll warning"""
        popup.destroy()  # Close the popup
        if stop:
            self.stop_export()

    def _show_scroll_warning(self):
        """Show scroll warning popup in a thread-safe way"""
        self.root.after(0, self.create_scroll_warning_popup)

    def _export_messages(self):
        """Export messages with improved file handling and GUI updates"""
        try:
            with self.driver_lock:
                if not self.driver:
                    return

                messages_container = None
                last_message = None  # Track last status message
                scroll_count = 0  # Track consecutive scroll messages

                # Wait for messages container with timeout and stop check
                start_time = time.time()
                while time.time() - start_time < 30 and self.export_running.is_set():
                    try:
                        messages_container = self.driver.find_element(
                            By.CSS_SELECTOR,
                            "div.x78zum5.xdt5ytf.x1iyjqo2"
                        )
                        if messages_container:
                            break
                    except:
                        time.sleep(0.5)

                if not messages_container or not self.export_running.is_set():
                    return

                output_file = self.output_path.get()
                os.makedirs(os.path.dirname(output_file), exist_ok=True)

                # Open file in write mode first to clear it, then reopen in append mode
                with open(output_file, "w", encoding="utf-8") as f:
                    # Write initial session marker
                    session_start = f"=== Export Session Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n"
                    f.write(session_start)
                    f.flush()
                    self.message_queue.put({
                        'type': 'status',
                        'message': session_start.strip(),
                        'level': 'info'
                    })

                # Reopen in append mode for the rest of the export
                with open(output_file, "a", encoding="utf-8") as f:
                    message_count = 0
                    last_height = 0
                    no_progress_count = 0
                    message_xpath = ".//div/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div/span/div/div/div/span/div"
                    sender_xpath = "./ancestor::div[contains(@class, 'x1n2onr6')]/div[1]/div/div/h4/div/div/span/span/span"
                    reply_sender_xpath = "./ancestor::div[contains(@class, 'x1n2onr6')]/div[1]/div/div[1]/div/div/h4/div/div/div/div[2]/span/span"
                    wait = WebDriverWait(self.driver, 10)

                    while self.export_running.is_set():
                        try:
                            visible_messages = messages_container.find_elements(By.XPATH, message_xpath)

                            if not visible_messages:
                                time.sleep(1)
                                continue

                            total_visible = len(visible_messages)
                            processed_in_view = 0

                            for message in visible_messages[::-1]:
                                if not self.export_running.is_set():
                                    break

                                message_id = self.get_message_identifier(message)
                                if message_id and message_id in self.processed_messages:
                                    processed_in_view += 1
                                    continue

                                try:
                                    # Center current message
                                    actions = ActionChains(self.driver)
                                    actions.move_to_element(message).perform()
                                except:
                                    try:
                                        self.driver.execute_script(
                                            "arguments[0].scrollIntoView({block: 'center'});",
                                            message
                                        )
                                    except:
                                        continue

                                # Reset scroll count when processing messages
                                scroll_count = 0

                                # Skip pure image messages
                                has_images = message.find_elements(By.TAG_NAME, "img")
                                has_image_text = "image" in message.get_attribute("innerHTML").lower()
                                if has_images and has_image_text and not message.text:
                                    if message_id:
                                        self.processed_messages.add(message_id)
                                    processed_in_view += 1
                                    continue

                                # Get message content
                                try:
                                    sender_name = ""
                                    if self.chat_type.get() == "group":
                                        try:
                                            sender_element = message.find_element(By.XPATH, sender_xpath)
                                            sender_name = sender_element.text.strip()
                                        except:
                                            try:
                                                sender_element = message.find_element(By.XPATH, reply_sender_xpath)
                                                sender_name = sender_element.text.strip()
                                            except:
                                                sender_name = ""

                                    color = self._get_message_color(message)
                                    content = self.driver.execute_script("""
                                        function extractContent(element) {
                                            var text = '';
                                            function processNode(node) {
                                                if (node.nodeType === Node.TEXT_NODE) {
                                                    text += node.textContent;
                                                } else if (node.tagName === 'A') {
                                                    text += node.textContent + ' [' + node.href + '] ';
                                                } else {
                                                    for (var i = 0; i < node.childNodes.length; i++) {
                                                        processNode(node.childNodes[i]);
                                                    }
                                                }
                                            }
                                            processNode(element);
                                            return text;
                                        }
                                        return extractContent(arguments[0]);
                                    """, message)

                                    content = content.strip()
                                    if content:
                                        message_count += 1
                                        formatted_message = ""

                                        if self.chat_type.get() != "group":
                                            color_name = self.rgb_to_color_name(color)
                                            if color_name == 'Azure':
                                                color_name = 'You'
                                            if color_name == 'White':
                                                color_name = 'Them'
                                            formatted_message = f"[{color_name}] {sender_name}{content}"
                                            f.write(formatted_message + "\n")
                                            f.flush()
                                        else:
                                            color_name = self.rgb_to_color_name(color)
                                            if color_name == 'Azure':
                                                sender_name = 'You'
                                            if sender_name != "":
                                                formatted_message = f"{content}\n[{sender_name}]"
                                                f.write(formatted_message + "\n")
                                            else:
                                                formatted_message = content
                                                f.write(formatted_message + "\n")
                                            f.flush()

                                        # Update GUI with the message
                                        self.message_queue.put({
                                            'type': 'status',
                                            'message': formatted_message,
                                            'level': 'info'
                                        })
                                        last_message = formatted_message

                                        if message_id:
                                            self.processed_messages.add(message_id)

                                        # Update progress every 10 messages
                                        if message_count % 10 == 0:
                                            progress_msg = f"Processed {message_count} messages..."
                                            self.message_queue.put({
                                                'type': 'status',
                                                'message': progress_msg,
                                                'level': 'info'
                                            })
                                            last_message = progress_msg

                                except Exception as e:
                                    self.message_queue.put({
                                        'type': 'status',
                                        'message': f'Error processing message: {str(e)}',
                                        'level': 'warning'
                                    })
                                    continue

                                processed_in_view += 1

                            # Scroll handling
                            if processed_in_view >= total_visible and self.export_running.is_set():
                                try:
                                    first_message = visible_messages[0]
                                    actions = ActionChains(self.driver)
                                    actions.move_to_element(first_message).perform()
                                    actions.send_keys(Keys.PAGE_UP).perform()
                                    time.sleep(2)

                                    # Check if last message was also a scroll message
                                    if last_message == 'Scrolling to load more messages...':
                                        scroll_count += 1
                                    else:
                                        scroll_count = 0

                                    # Update GUI about scrolling
                                    self.message_queue.put({
                                        'type': 'status',
                                        'message': 'Scrolling to load more messages...',
                                        'level': 'info'
                                    })
                                    last_message = 'Scrolling to load more messages...'

                                    # Show warning popup if multiple scroll attempts
                                    if scroll_count >= 1:
                                        self._show_scroll_warning()
                                        # Reset scroll count after showing warning
                                        scroll_count = 0
                                except:
                                    try:
                                        self.driver.execute_script("window.scrollBy(0, -1000);")
                                        time.sleep(2)
                                    except:
                                        pass

                        except Exception as e:
                            self.message_queue.put({
                                'type': 'status',
                                'message': f'Error in main loop: {str(e)}',
                                'level': 'warning'
                            })
                            time.sleep(1)
                            continue

                    # Write session end marker
                    session_end = f"\n=== Export Session Ended: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n"
                    f.write(session_end)
                    f.flush()
                    self.message_queue.put({
                        'type': 'status',
                        'message': session_end.strip(),
                        'level': 'info'
                    })

                    # Final summary
                    summary_msg = f"Export completed! Saved {message_count} messages to {output_file}"
                    self.message_queue.put({
                        'type': 'status',
                        'message': summary_msg,
                        'level': 'info'
                    })

        except Exception as e:
            self.message_queue.put({
                'type': 'status',
                'message': f'Critical export error: {str(e)}',
                'level': 'error'
            })
        finally:
            if not self.export_running.is_set():
                self._cleanup()

    def get_message_color(self, message):
        """Get the background color of a message"""
        try:
            bubble = message.find_element(By.XPATH, ".//div[contains(@class, 'x1n2onr6')]")
            color = self.driver.execute_script(
                "return window.getComputedStyle(arguments[0]).backgroundColor;",
                bubble
            )
            return color
        except:
            return None

    def get_message_content(self, message):
        """Extract message content with improved error handling"""
        try:
            content = self.driver.execute_script("""
                function extractContent(element) {
                    var text = '';
                    function processNode(node) {
                        if (node.nodeType === Node.TEXT_NODE) {
                            text += node.textContent;
                        } else if (node.tagName === 'A') {
                            text += node.textContent + ' [' + node.href + '] ';
                        } else if (node.tagName === 'IMG') {
                            text += '[Image] ';
                        } else {
                            for (var i = 0; i < node.childNodes.length; i++) {
                                processNode(node.childNodes[i]);
                            }
                        }
                        return text;
                    }
                    return processNode(element);
                }
                return extractContent(arguments[0]);
            """, message)
            return content.strip()
        except:
            try:
                return message.text.strip()
            except:
                return ""

    def rgb_to_color_name(self, rgb_string):
        """Convert RGB color to descriptive name with support for 20+ distinct colors

        Args:
            rgb_string (str): RGB color in format 'rgb(r,g,b)' or 'rgba(r,g,b,a)'

        Returns:
            str: Detailed color name description
        """
        try:
            # Parse RGB values
            rgb_values = rgb_string.strip('rgba()').split(',')
            r = int(rgb_values[0])
            g = int(rgb_values[1])
            b = int(rgb_values[2])

            # Check for grayscale first
            if max(abs(r - g), abs(r - b), abs(g - b)) < 20:
                if r < 30:
                    return "Black"
                elif r > 225:
                    return "White"
                elif r > 160:
                    return "Light Gray"
                elif r > 90:
                    return "Gray"
                else:
                    return "Dark Gray"

            # Calculate color dominance
            max_val = max(r, g, b)
            min_val = min(r, g, b)
            delta = max_val - min_val

            # Calculate hue
            if delta == 0:
                hue = 0
            elif max_val == r:
                hue = 60 * (((g - b) / delta) % 6)
            elif max_val == g:
                hue = 60 * (((b - r) / delta) + 2)
            else:
                hue = 60 * (((r - g) / delta) + 4)

            # Calculate saturation
            saturation = 0 if max_val == 0 else (delta / max_val) * 100

            # Calculate value/brightness
            value = (max_val / 255) * 100

            # Early returns for special cases
            if saturation < 10:
                return "Gray"
            if value < 15:
                return "Black"
            elif value > 95 and saturation < 20:
                return "White"

            # Detailed color determination based on hue ranges
            if 0 <= hue <= 10 or hue >= 350:
                if value > 80:
                    return "Pink" if saturation < 50 else "Red"
                elif value > 50:
                    return "Dark Red"
                else:
                    return "Maroon"

            elif 10 < hue <= 20:
                if value > 80:
                    return "Salmon"
                else:
                    return "Dark Salmon"

            elif 20 < hue <= 40:
                if value > 80:
                    return "Orange"
                elif value > 50:
                    return "Dark Orange"
                else:
                    return "Brown"

            elif 40 < hue <= 70:
                if value > 80:
                    return "Yellow"
                elif value > 50:
                    return "Gold"
                else:
                    return "Olive"

            elif 70 < hue <= 100:
                if value > 80:
                    return "Lime"
                elif value > 50:
                    return "Green"
                else:
                    return "Dark Green"

            elif 100 < hue <= 140:
                if value > 80:
                    return "Light Green"
                elif value > 50:
                    return "Sea Green"
                else:
                    return "Forest Green"

            elif 140 < hue <= 170:
                if value > 80:
                    return "Aqua"
                elif value > 50:
                    return "Teal"
                else:
                    return "Dark Teal"

            elif 170 < hue <= 200:
                if value > 80:
                    return "Light Blue"
                elif value > 50:
                    return "Sky Blue"
                else:
                    return "Steel Blue"

            elif 200 < hue <= 240:
                if value > 80:
                    return "Azure"
                elif value > 50:
                    return "Blue"
                else:
                    return "Navy"

            elif 240 < hue <= 280:
                if value > 80:
                    return "Violet"
                elif value > 50:
                    return "Purple"
                else:
                    return "Dark Purple"

            elif 280 < hue <= 320:
                if value > 80:
                    return "Magenta"
                elif value > 50:
                    return "Hot Pink"
                else:
                    return "Deep Pink"

            elif 320 < hue < 350:
                if value > 80:
                    return "Light Pink"
                elif value > 50:
                    return "Rose"
                else:
                    return "Dark Rose"

            return "Unknown"

        except Exception as e:
            return "Unknown"

    def get_sender_name(self, message, color):
        """Get sender name with improved error handling"""
        try:
            sender_name = ""
            xpath_expressions = [
                "./ancestor::div[contains(@class, 'x1n2onr6')]/div[1]/div/div/h4/div/div/span/span/span",
                ".//div[contains(@class, 'x1n2onr6')]//h4//span/span/span",
                ".//h4//span/span/span"
            ]

            for xpath in xpath_expressions:
                try:
                    sender_element = message.find_element(By.XPATH, xpath)
                    sender_name = sender_element.text.strip()
                    if sender_name:
                        break
                except:
                    continue

            # Identify if message is from "Me"
            color_name = self.rgb_to_color_name(color)
            if color_name == 'Blue':
                sender_name = 'Me'

            return sender_name
        except:
            return ""

    def run(self):
        """Start the application"""
        self.root.mainloop()

def main():
    def handle_exception(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return

        error_msg = f"An unexpected error occurred: {str(exc_value)}"
        try:
            app.update_status(error_msg, "error")
        except:
            print(error_msg)

    sys.excepthook = handle_exception

    app = ModernMessengerExporter()
    app.run()


if __name__ == "__main__":
    main()