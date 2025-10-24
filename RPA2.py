import sys
import json
import time
import threading
from datetime import datetime
from pathlib import Path
import base64
from io import BytesIO

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QListWidget, QComboBox,
    QGroupBox, QTextEdit, QTabWidget, QMessageBox, QSplitter,
    QListWidgetItem, QSpinBox, QCheckBox, QDialog, QDialogButtonBox,
    QFormLayout, QScrollArea, QFrame
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QSize
from PySide6.QtGui import QIcon, QFont, QColor, QPixmap, QImage
from functools import partial

import pyautogui
from pynput import mouse, keyboard
from PIL import Image

import pandas as pd
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

class ActionDialog(QDialog):
    """Dialog for adding/editing advanced actions"""
    def __init__(self, parent=None, action_type="click"):
        super().__init__(parent)
        self.setWindowTitle("Add Advanced Action")
        self.setModal(True)
        self.action_type = action_type
        self.init_ui()
 
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        form_layout = QFormLayout()
        
        # Action type selector
        self.type_combo = QComboBox()
        self.type_combo.addItems([
            "click", "mouse_move", "key", "type_text", 
            "wait", "screenshot", "if_condition", "loop",
            "web_navigate", "web_click", "web_type", "web_extract",  # NEW
            "excel_read", "excel_write", "csv_read", "csv_write"  # NEW
        ])   
        self.type_combo.setCurrentText(self.action_type)
        self.type_combo.currentTextChanged.connect(self.on_type_changed)
        form_layout.addRow("Action Type:", self.type_combo)
        
        # Dynamic fields container
        self.fields_widget = QWidget()
        self.fields_layout = QFormLayout(self.fields_widget)
        
        self.update_fields()
        
        layout.addLayout(form_layout)
        layout.addWidget(self.fields_widget)
        
        # Buttons
        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
    def on_type_changed(self, action_type):
        self.action_type = action_type
        self.update_fields()
        
    def update_fields(self):
        # Clear existing fields
        while self.fields_layout.count():
            child = self.fields_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
                
        # Add fields based on action type
        if self.action_type == "click":
            self.x_spin = QSpinBox()
            self.x_spin.setRange(0, 10000)
            self.fields_layout.addRow("X:", self.x_spin)
            
            self.y_spin = QSpinBox()
            self.y_spin.setRange(0, 10000)
            self.fields_layout.addRow("Y:", self.y_spin)
            
            self.button_combo = QComboBox()
            self.button_combo.addItems(["left", "right", "middle"])
            self.fields_layout.addRow("Button:", self.button_combo)
            
        elif self.action_type == "mouse_move":
            self.x_spin = QSpinBox()
            self.x_spin.setRange(0, 10000)
            self.fields_layout.addRow("X:", self.x_spin)
            
            self.y_spin = QSpinBox()
            self.y_spin.setRange(0, 10000)
            self.fields_layout.addRow("Y:", self.y_spin)
            
            self.duration_spin = QSpinBox()
            self.duration_spin.setRange(0, 10)
            self.duration_spin.setValue(1)
            self.fields_layout.addRow("Duration (s):", self.duration_spin)
            
        elif self.action_type == "key":
            self.key_input = QLineEdit()
            self.key_input.setPlaceholderText("e.g., enter, space, ctrl")
            self.fields_layout.addRow("Key:", self.key_input)
            
        elif self.action_type == "type_text":
            self.text_input = QLineEdit()
            self.text_input.setPlaceholderText("Text to type...")
            self.fields_layout.addRow("Text:", self.text_input)
            
            self.interval_spin = QSpinBox()
            self.interval_spin.setRange(0, 1000)
            self.interval_spin.setValue(10)
            self.fields_layout.addRow("Interval (ms):", self.interval_spin)
            
        elif self.action_type == "wait":
            self.wait_spin = QSpinBox()
            self.wait_spin.setRange(1, 3600)
            self.wait_spin.setValue(2)
            self.fields_layout.addRow("Wait time (s):", self.wait_spin)
            
        elif self.action_type == "if_condition":
            self.condition_input = QLineEdit()
            self.condition_input.setPlaceholderText("e.g., pixel_color(100,100) == (255,0,0)")
            self.fields_layout.addRow("Condition:", self.condition_input)
            
        elif self.action_type == "loop":
            self.loop_count = QSpinBox()
            self.loop_count.setRange(1, 1000)
            self.loop_count.setValue(5)
            self.fields_layout.addRow("Iterations:", self.loop_count)

        elif self.action_type == "web_navigate":
            self.url_input = QLineEdit()
            self.url_input.setPlaceholderText("https://example.com")
            self.fields_layout.addRow("URL:", self.url_input)
            
        elif self.action_type == "web_click":
            self.selector_type = QComboBox()
            self.selector_type.addItems(["id", "name", "xpath", "css", "class"])
            self.fields_layout.addRow("Selector Type:", self.selector_type)
            
            self.selector_input = QLineEdit()
            self.selector_input.setPlaceholderText("Element selector...")
            self.fields_layout.addRow("Selector:", self.selector_input)
            
        elif self.action_type == "web_type":
            self.selector_type = QComboBox()
            self.selector_type.addItems(["id", "name", "xpath", "css", "class"])
            self.fields_layout.addRow("Selector Type:", self.selector_type)
            
            self.selector_input = QLineEdit()
            self.selector_input.setPlaceholderText("Element selector...")
            self.fields_layout.addRow("Selector:", self.selector_input)
            
            self.web_text_input = QLineEdit()
            self.web_text_input.setPlaceholderText("Text to type...")
            self.fields_layout.addRow("Text:", self.web_text_input)
            
        elif self.action_type == "web_extract":
            self.selector_type = QComboBox()
            self.selector_type.addItems(["id", "name", "xpath", "css", "class"])
            self.fields_layout.addRow("Selector Type:", self.selector_type)
            
            self.selector_input = QLineEdit()
            self.selector_input.setPlaceholderText("Element selector...")
            self.fields_layout.addRow("Selector:", self.selector_input)
            
            self.var_name_input = QLineEdit()
            self.var_name_input.setPlaceholderText("Variable name to store data")
            self.fields_layout.addRow("Save to Variable:", self.var_name_input)
            
        elif self.action_type == "excel_read":
            self.file_path_input = QLineEdit()
            self.file_path_input.setPlaceholderText("path/to/file.xlsx")
            self.fields_layout.addRow("File Path:", self.file_path_input)
            
            self.sheet_name_input = QLineEdit()
            self.sheet_name_input.setPlaceholderText("Sheet1")
            self.fields_layout.addRow("Sheet Name:", self.sheet_name_input)
            
            self.var_name_input = QLineEdit()
            self.var_name_input.setPlaceholderText("Variable name")
            self.fields_layout.addRow("Save to Variable:", self.var_name_input)
            
        elif self.action_type == "excel_write":
            self.file_path_input = QLineEdit()
            self.file_path_input.setPlaceholderText("path/to/file.xlsx")
            self.fields_layout.addRow("File Path:", self.file_path_input)
            
            self.sheet_name_input = QLineEdit()
            self.sheet_name_input.setPlaceholderText("Sheet1")
            self.fields_layout.addRow("Sheet Name:", self.sheet_name_input)
            
            self.data_var_input = QLineEdit()
            self.data_var_input.setPlaceholderText("Variable containing data")
            self.fields_layout.addRow("Data Variable:", self.data_var_input)
            
        elif self.action_type == "csv_read":
            self.file_path_input = QLineEdit()
            self.file_path_input.setPlaceholderText("path/to/file.csv")
            self.fields_layout.addRow("File Path:", self.file_path_input)
            
            self.var_name_input = QLineEdit()
            self.var_name_input.setPlaceholderText("Variable name")
            self.fields_layout.addRow("Save to Variable:", self.var_name_input)
            
        elif self.action_type == "csv_write":
            self.file_path_input = QLineEdit()
            self.file_path_input.setPlaceholderText("path/to/file.csv")
            self.fields_layout.addRow("File Path:", self.file_path_input)
            
            self.data_var_input = QLineEdit()
            self.data_var_input.setPlaceholderText("Variable containing data")
            self.fields_layout.addRow("Data Variable:", self.data_var_input)

            
    def get_action(self):
        """Return the configured action"""
        action = {
            "type": self.action_type,
            "timestamp": datetime.now().isoformat()
        }
        
        if self.action_type == "click":
            action["x"] = self.x_spin.value()
            action["y"] = self.y_spin.value()
            action["button"] = self.button_combo.currentText()
            
        elif self.action_type == "mouse_move":
            action["x"] = self.x_spin.value()
            action["y"] = self.y_spin.value()
            action["duration"] = self.duration_spin.value()
            
        elif self.action_type == "key":
            action["key"] = self.key_input.text()
            
        elif self.action_type == "type_text":
            action["text"] = self.text_input.text()
            action["interval"] = self.interval_spin.value() / 1000.0
            
        elif self.action_type == "wait":
            action["duration"] = self.wait_spin.value()
            
        elif self.action_type == "if_condition":
            action["condition"] = self.condition_input.text()
            
        elif self.action_type == "loop":
            action["iterations"] = self.loop_count.value()
            action["actions"] = []  # Nested actions

        elif self.action_type == "web_navigate":
            action["url"] = self.url_input.text()
            
        elif self.action_type == "web_click":
            action["selector_type"] = self.selector_type.currentText()
            action["selector"] = self.selector_input.text()
            
        elif self.action_type == "web_type":
            action["selector_type"] = self.selector_type.currentText()
            action["selector"] = self.selector_input.text()
            action["text"] = self.web_text_input.text()
            
        elif self.action_type == "web_extract":
            action["selector_type"] = self.selector_type.currentText()
            action["selector"] = self.selector_input.text()
            action["variable"] = self.var_name_input.text()
            
        elif self.action_type == "excel_read":
            action["file_path"] = self.file_path_input.text()
            action["sheet_name"] = self.sheet_name_input.text()
            action["variable"] = self.var_name_input.text()
            
        elif self.action_type == "excel_write":
            action["file_path"] = self.file_path_input.text()
            action["sheet_name"] = self.sheet_name_input.text()
            action["data_variable"] = self.data_var_input.text()
            
        elif self.action_type == "csv_read":
            action["file_path"] = self.file_path_input.text()
            action["variable"] = self.var_name_input.text()
            
        elif self.action_type == "csv_write":
            action["file_path"] = self.file_path_input.text()
            action["data_variable"] = self.data_var_input.text()
            
        return action


class RecorderThread(QThread):
    """Thread for recording user actions"""
    action_recorded = Signal(dict)
    
    def __init__(self):
        super().__init__()
        self.recording = False
        self.capture_movement = False
        self.capture_screenshots = False
        self.mouse_listener = None
        self.keyboard_listener = None
        self.last_move_time = 0
        
    def run(self):
        def on_click(x, y, button, pressed):
            if self.recording and pressed:
                action = {
                    "type": "click",
                    "x": x,
                    "y": y,
                    "button": str(button).split('.')[-1],
                    "timestamp": datetime.now().isoformat()
                }
                
                # Capture screenshot if enabled
                if self.capture_screenshots:
                    try:
                        screenshot = pyautogui.screenshot(region=(x-50, y-50, 100, 100))
                        buffer = BytesIO()
                        screenshot.save(buffer, format='PNG')
                        action["screenshot"] = base64.b64encode(buffer.getvalue()).decode()
                    except:
                        pass
                        
                self.action_recorded.emit(action)
                
        def on_move(x, y):
            if self.recording and self.capture_movement:
                current_time = time.time()
                # Only record movement every 0.5 seconds to avoid spam
                if current_time - self.last_move_time > 0.5:
                    action = {
                        "type": "mouse_move",
                        "x": x,
                        "y": y,
                        "duration": 0.5,
                        "timestamp": datetime.now().isoformat()
                    }
                    self.action_recorded.emit(action)
                    self.last_move_time = current_time
                
        def on_key(key):
            if self.recording:
                try:
                    key_str = key.char
                except AttributeError:
                    key_str = str(key).replace('Key.', '')
                    
                action = {
                    "type": "key",
                    "key": key_str,
                    "timestamp": datetime.now().isoformat()
                }
                self.action_recorded.emit(action)
                
        self.mouse_listener = mouse.Listener(on_click=on_click, on_move=on_move)
        self.keyboard_listener = keyboard.Listener(on_press=on_key)
        
        self.mouse_listener.start()
        self.keyboard_listener.start()
        
        self.mouse_listener.join()
        
    def start_recording(self, capture_movement=False, capture_screenshots=False):
        self.recording = True
        self.capture_movement = capture_movement
        self.capture_screenshots = capture_screenshots
        if not self.isRunning():
            self.start()
            
    def stop_recording(self):
        self.recording = False
        if self.mouse_listener:
            self.mouse_listener.stop()
        if self.keyboard_listener:
            self.keyboard_listener.stop()


class ExecutorThread(QThread):
    """Thread for executing workflows"""
    progress_update = Signal(str)
    execution_complete = Signal(bool, str)
    screenshot_captured = Signal(QPixmap)
    
    def __init__(self, workflow, delay=0.5):
        super().__init__()
        self.workflow = workflow
        self.delay = delay
        self.variables = {}
        self.driver = None
        
    def run(self):
        try:
            actions = self.workflow.get("actions", [])
            self.execute_actions(actions)
            self.execution_complete.emit(True, "Workflow completed successfully!")
            
        except Exception as e:
            self.execution_complete.emit(False, f"Error: {str(e)}")
        finally:
            self.cleanup_selenium()  # ADD THIS
            
    def execute_actions(self, actions, loop_count=1):
        """Execute a list of actions"""
        for iteration in range(loop_count):
            for i, action in enumerate(actions):
                action_type = action["type"]
                
                self.progress_update.emit(
                    f"Executing step {i+1}/{len(actions)}: {action_type}"
                )
                
                if action_type == "click":
                    pyautogui.click(action["x"], action["y"], button=action.get("button", "left"))
                    
                elif action_type == "mouse_move":
                    duration = action.get("duration", 0.5)
                    pyautogui.moveTo(action["x"], action["y"], duration=duration)
                    
                elif action_type == "key":
                    try:
                        pyautogui.press(action["key"])
                    except:
                        pyautogui.write(action["key"])
                        
                elif action_type == "type_text":
                    interval = action.get("interval", 0.01)
                    pyautogui.write(action["text"], interval=interval)
                    
                elif action_type == "wait":
                    wait_time = action.get("duration", 2)
                    self.progress_update.emit(f"Waiting {wait_time} seconds...")
                    time.sleep(wait_time)
                    
                elif action_type == "screenshot":
                    screenshot = pyautogui.screenshot()
                    # Convert PIL Image to QPixmap
                    screenshot.save("temp_screenshot.png")
                    pixmap = QPixmap("temp_screenshot.png")
                    self.screenshot_captured.emit(pixmap)
                    
                elif action_type == "if_condition":
                    # Simple condition evaluation (can be expanded)
                    condition = action.get("condition", "")
                    # For safety, we'll skip complex evaluation
                    self.progress_update.emit(f"Condition check: {condition}")
                    
                elif action_type == "loop":
                    iterations = action.get("iterations", 1)
                    nested_actions = action.get("actions", [])
                    self.progress_update.emit(f"Starting loop: {iterations} iterations")
                    self.execute_actions(nested_actions, iterations)

                elif action_type == "web_navigate":
                    self.setup_selenium()
                    if self.driver:
                        url = action.get("url", "")
                        self.progress_update.emit(f"Navigating to {url}")
                        self.driver.get(url)
                        
                elif action_type == "web_click":
                    if self.driver:
                        selector_type = action.get("selector_type", "id")
                        selector = action.get("selector", "")
                        
                        by_type = {
                            "id": By.ID,
                            "name": By.NAME,
                            "xpath": By.XPATH,
                            "css": By.CSS_SELECTOR,
                            "class": By.CLASS_NAME
                        }.get(selector_type, By.ID)
                        
                        try:
                            element = WebDriverWait(self.driver, 10).until(
                                EC.element_to_be_clickable((by_type, selector))
                            )
                            element.click()
                            self.progress_update.emit(f"Clicked element: {selector}")
                        except TimeoutException:
                            self.progress_update.emit(f"Element not found: {selector}")
                            
                elif action_type == "web_type":
                    if self.driver:
                        selector_type = action.get("selector_type", "id")
                        selector = action.get("selector", "")
                        text = action.get("text", "")
                        
                        by_type = {
                            "id": By.ID,
                            "name": By.NAME,
                            "xpath": By.XPATH,
                            "css": By.CSS_SELECTOR,
                            "class": By.CLASS_NAME
                        }.get(selector_type, By.ID)
                        
                        try:
                            element = WebDriverWait(self.driver, 10).until(
                                EC.presence_of_element_located((by_type, selector))
                            )
                            element.clear()
                            element.send_keys(text)
                            self.progress_update.emit(f"Typed into element: {selector}")
                        except TimeoutException:
                            self.progress_update.emit(f"Element not found: {selector}")
                            
                elif action_type == "web_extract":
                    if self.driver:
                        selector_type = action.get("selector_type", "id")
                        selector = action.get("selector", "")
                        var_name = action.get("variable", "extracted_data")
                        
                        by_type = {
                            "id": By.ID,
                            "name": By.NAME,
                            "xpath": By.XPATH,
                            "css": By.CSS_SELECTOR,
                            "class": By.CLASS_NAME
                        }.get(selector_type, By.ID)
                        
                        try:
                            element = WebDriverWait(self.driver, 10).until(
                                EC.presence_of_element_located((by_type, selector))
                            )
                            self.variables[var_name] = element.text
                            self.progress_update.emit(f"Extracted data to variable: {var_name}")
                        except TimeoutException:
                            self.progress_update.emit(f"Element not found: {selector}")
                            
                elif action_type == "excel_read":
                    file_path = action.get("file_path", "")
                    sheet_name = action.get("sheet_name", "Sheet1")
                    var_name = action.get("variable", "excel_data")
                    
                    try:
                        df = pd.read_excel(file_path, sheet_name=sheet_name)
                        self.variables[var_name] = df
                        self.progress_update.emit(f"Read Excel file: {file_path}")
                    except Exception as e:
                        self.progress_update.emit(f"Error reading Excel: {str(e)}")
                        
                elif action_type == "excel_write":
                    file_path = action.get("file_path", "")
                    sheet_name = action.get("sheet_name", "Sheet1")
                    data_var = action.get("data_variable", "")
                    
                    if data_var in self.variables:
                        try:
                            df = self.variables[data_var]
                            df.to_excel(file_path, sheet_name=sheet_name, index=False)
                            self.progress_update.emit(f"Wrote Excel file: {file_path}")
                        except Exception as e:
                            self.progress_update.emit(f"Error writing Excel: {str(e)}")
                    else:
                        self.progress_update.emit(f"Variable not found: {data_var}")
                        
                elif action_type == "csv_read":
                    file_path = action.get("file_path", "")
                    var_name = action.get("variable", "csv_data")
                    
                    try:
                        df = pd.read_csv(file_path)
                        self.variables[var_name] = df
                        self.progress_update.emit(f"Read CSV file: {file_path}")
                    except Exception as e:
                        self.progress_update.emit(f"Error reading CSV: {str(e)}")
                        
                elif action_type == "csv_write":
                    file_path = action.get("file_path", "")
                    data_var = action.get("data_variable", "")
                    
                    if data_var in self.variables:
                        try:
                            df = self.variables[data_var]
                            df.to_csv(file_path, index=False)
                            self.progress_update.emit(f"Wrote CSV file: {file_path}")
                        except Exception as e:
                            self.progress_update.emit(f"Error writing CSV: {str(e)}")
                    else:
                        self.progress_update.emit(f"Variable not found: {data_var}")
 
                time.sleep(self.delay)
                
    def setup_selenium(self):
        """Initialize Selenium WebDriver"""
        if not self.driver:
            try:
                self.driver = webdriver.Chrome()  # or webdriver.Firefox()
            except:
                # Fallback to other browsers or show error
                pass
                
    def cleanup_selenium(self):
        """Close Selenium WebDriver"""
        if self.driver:
            self.driver.quit()
            self.driver = None

class RPAMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.workflows = []
        self.current_actions = []
        self.recorder_thread = RecorderThread()
        self.executor_thread = None
        self.workflows_file = Path("workflows.json")
        
        self.init_ui()
        self.load_workflows()
        
        self.recorder_thread.action_recorded.connect(self.on_action_recorded)
        
    def init_ui(self):
        self.setWindowTitle("Advanced RPA Automation Tool")
        self.setGeometry(100, 100, 1200, 800)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        
        # Title
        title_label = QLabel("🤖 Advanced RPA Automation Tool")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # Tab widget
        tabs = QTabWidget()
        main_layout.addWidget(tabs)
        
        tabs.addTab(self.create_record_tab(), "📹 Record")
        tabs.addTab(self.create_execute_tab(), "▶️ Execute")
        tabs.addTab(self.create_manage_tab(), "📋 Manage")
        tabs.addTab(self.create_advanced_tab(), "⚙️ Advanced")
        tabs.addTab(self.create_web_data_tab(), "🌐 Web & Data")
        
        self.statusBar().showMessage("Ready")

    def create_web_data_tab(self):
        """Create web and data automation tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Web Automation
        web_group = QGroupBox("Web Automation (Selenium)")
        web_layout = QVBoxLayout()
        web_layout.addWidget(QLabel("Automate web browsers:"))
        web_layout.addWidget(QLabel("• Navigate to URLs"))
        web_layout.addWidget(QLabel("• Click elements"))
        web_layout.addWidget(QLabel("• Fill forms"))
        web_layout.addWidget(QLabel("• Extract data"))
        
        web_buttons = QHBoxLayout()
        add_navigate_btn = QPushButton("Navigate URL")
        add_navigate_btn.clicked.connect(partial(self.add_manual_action, "web_navigate"))
        web_buttons.addWidget(add_navigate_btn)
        
        add_web_click_btn = QPushButton("Web Click")
        add_web_click_btn.clicked.connect(partial(self.add_manual_action, "web_click"))
        web_buttons.addWidget(add_web_click_btn)
        
        add_web_type_btn = QPushButton("Web Type")
        add_web_type_btn.clicked.connect(partial(self.add_manual_action, "web_type"))
        web_buttons.addWidget(add_web_type_btn)
        
        add_web_extract_btn = QPushButton("Extract Data")
        add_web_extract_btn.clicked.connect(partial(self.add_manual_action, "web_extract"))
        web_buttons.addWidget(add_web_extract_btn)
        
        web_layout.addLayout(web_buttons)
        web_group.setLayout(web_layout)
        layout.addWidget(web_group)
        
        # Excel Automation
        excel_group = QGroupBox("Excel Automation")
        excel_layout = QVBoxLayout()
        excel_layout.addWidget(QLabel("Process Excel files:"))
        excel_layout.addWidget(QLabel("• Read data from Excel"))
        excel_layout.addWidget(QLabel("• Write data to Excel"))
        excel_layout.addWidget(QLabel("• Process multiple sheets"))
        
        excel_buttons = QHBoxLayout()
        add_excel_read_btn = QPushButton("Read Excel")
        add_excel_read_btn.clicked.connect(partial(self.add_manual_action, "excel_read"))
        excel_buttons.addWidget(add_excel_read_btn)
        
        add_excel_write_btn = QPushButton("Write Excel")
        add_excel_write_btn.clicked.connect(partial(self.add_manual_action, "excel_write"))
        excel_buttons.addWidget(add_excel_write_btn)
        
        excel_layout.addLayout(excel_buttons)
        excel_group.setLayout(excel_layout)
        layout.addWidget(excel_group)
        
        # CSV Automation
        csv_group = QGroupBox("CSV Automation")
        csv_layout = QVBoxLayout()
        csv_layout.addWidget(QLabel("Process CSV files:"))
        csv_layout.addWidget(QLabel("• Read CSV data"))
        csv_layout.addWidget(QLabel("• Write CSV files"))
        csv_layout.addWidget(QLabel("• Data transformation"))
        
        csv_buttons = QHBoxLayout()
        add_csv_read_btn = QPushButton("Read CSV")
        add_csv_read_btn.clicked.connect(partial(self.add_manual_action, "csv_read"))
        csv_buttons.addWidget(add_csv_read_btn)
        
        add_csv_write_btn = QPushButton("Write CSV")
        add_csv_write_btn.clicked.connect(partial(self.add_manual_action, "csv_write"))
        csv_buttons.addWidget(add_csv_write_btn)
        
        csv_layout.addLayout(csv_buttons)
        csv_group.setLayout(csv_layout)
        layout.addWidget(csv_group)
        
        layout.addStretch()
        
        # Tips
        tips_group = QGroupBox("💡 Tips")
        tips_layout = QVBoxLayout()
        tips = QTextEdit()
        tips.setReadOnly(True)
        tips.setMaximumHeight(120)
        tips.setHtml("""
        <ul>
            <li><b>Selenium:</b> Requires ChromeDriver or GeckoDriver installed</li>
            <li><b>Excel:</b> Use variable names to pass data between actions</li>
            <li><b>CSV:</b> Lightweight format for data processing</li>
            <li><b>Variables:</b> Store extracted data for later use</li>
        </ul>
        """)
        tips_layout.addWidget(tips)
        tips_group.setLayout(tips_layout)
        layout.addWidget(tips_group)
        
        return widget    
    
        
    def create_record_tab(self):
        """Create the recording tab with advanced options"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Recording options
        options_group = QGroupBox("Recording Options")
        options_layout = QVBoxLayout()
        
        self.capture_movement_check = QCheckBox("Capture mouse movements")
        self.capture_movement_check.setToolTip("Record mouse movements (may create many actions)")
        options_layout.addWidget(self.capture_movement_check)
        
        self.capture_screenshots_check = QCheckBox("Capture screenshots on click")
        self.capture_screenshots_check.setToolTip("Take screenshots at click locations for visual reference")
        options_layout.addWidget(self.capture_screenshots_check)
        
        options_group.setLayout(options_layout)
        layout.addWidget(options_group)
        
        # Recording controls
        controls_group = QGroupBox("Recording Controls")
        controls_layout = QVBoxLayout()
        
        buttons_layout = QHBoxLayout()
        
        self.start_record_btn = QPushButton("🔴 Start Recording")
        self.start_record_btn.setStyleSheet("background-color: #4CAF50; color: white; font-size: 14px; padding: 10px;")
        self.start_record_btn.clicked.connect(self.start_recording)
        
        self.stop_record_btn = QPushButton("⏹️ Stop Recording")
        self.stop_record_btn.setStyleSheet("background-color: #f44336; color: white; font-size: 14px; padding: 10px;")
        self.stop_record_btn.setEnabled(False)
        self.stop_record_btn.clicked.connect(self.stop_recording)
        
        buttons_layout.addWidget(self.start_record_btn)
        buttons_layout.addWidget(self.stop_record_btn)
        controls_layout.addLayout(buttons_layout)
        
        self.record_status_label = QLabel("Status: Ready to record")
        self.record_status_label.setStyleSheet("font-size: 13px; padding: 5px;")
        controls_layout.addWidget(self.record_status_label)
        
        controls_group.setLayout(controls_layout)
        layout.addWidget(controls_group)
        
        # Actions list
        actions_group = QGroupBox("Recorded Actions")
        actions_layout = QVBoxLayout()
        
        self.actions_list = QListWidget()
        self.actions_list.setStyleSheet("font-family: monospace;")
        actions_layout.addWidget(self.actions_list)
        
        # Action management buttons
        action_buttons = QHBoxLayout()
        
        add_action_btn = QPushButton("➕ Add Manual Action")
        add_action_btn.clicked.connect(self.add_manual_action)
        action_buttons.addWidget(add_action_btn)
        
        edit_action_btn = QPushButton("✏️ Edit Selected")
        edit_action_btn.clicked.connect(self.edit_action)
        action_buttons.addWidget(edit_action_btn)
        
        delete_action_btn = QPushButton("🗑️ Delete Selected")
        delete_action_btn.clicked.connect(self.delete_action)
        action_buttons.addWidget(delete_action_btn)
        
        clear_btn = QPushButton("Clear All")
        clear_btn.clicked.connect(self.clear_actions)
        action_buttons.addWidget(clear_btn)
        
        actions_layout.addLayout(action_buttons)
        actions_group.setLayout(actions_layout)
        layout.addWidget(actions_group)
        
        # Save workflow
        save_group = QGroupBox("Save Workflow")
        save_layout = QHBoxLayout()
        
        self.workflow_name_input = QLineEdit()
        self.workflow_name_input.setPlaceholderText("Enter workflow name...")
        save_layout.addWidget(self.workflow_name_input)
        
        save_btn = QPushButton("💾 Save Workflow")
        save_btn.setStyleSheet("background-color: #2196F3; color: white; padding: 8px;")
        save_btn.clicked.connect(self.save_workflow)
        save_layout.addWidget(save_btn)
        
        save_group.setLayout(save_layout)
        layout.addWidget(save_group)
        
        return widget
        
    def create_execute_tab(self):
        """Create the execution tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Workflow selection
        select_group = QGroupBox("Select Workflow")
        select_layout = QVBoxLayout()
        
        self.workflow_combo = QComboBox()
        self.workflow_combo.currentIndexChanged.connect(self.on_workflow_selected)
        select_layout.addWidget(self.workflow_combo)
        
        select_group.setLayout(select_layout)
        layout.addWidget(select_group)
        
        # Workflow details
        details_group = QGroupBox("Workflow Details")
        details_layout = QVBoxLayout()
        
        self.workflow_details = QTextEdit()
        self.workflow_details.setReadOnly(True)
        self.workflow_details.setMaximumHeight(150)
        details_layout.addWidget(self.workflow_details)
        
        details_group.setLayout(details_layout)
        layout.addWidget(details_group)
        
        # Execution settings
        settings_group = QGroupBox("Execution Settings")
        settings_layout = QVBoxLayout()
        
        delay_layout = QHBoxLayout()
        delay_layout.addWidget(QLabel("Delay between actions (seconds):"))
        self.delay_spinbox = QSpinBox()
        self.delay_spinbox.setMinimum(0)
        self.delay_spinbox.setMaximum(10)
        self.delay_spinbox.setValue(1)
        delay_layout.addWidget(self.delay_spinbox)
        delay_layout.addStretch()
        settings_layout.addLayout(delay_layout)
        
        # Loop execution
        loop_layout = QHBoxLayout()
        self.loop_check = QCheckBox("Repeat workflow")
        loop_layout.addWidget(self.loop_check)
        loop_layout.addWidget(QLabel("times:"))
        self.loop_count = QSpinBox()
        self.loop_count.setMinimum(1)
        self.loop_count.setMaximum(1000)
        self.loop_count.setValue(1)
        self.loop_count.setEnabled(False)
        self.loop_check.toggled.connect(lambda checked: self.loop_count.setEnabled(checked))
        loop_layout.addWidget(self.loop_count)
        loop_layout.addStretch()
        settings_layout.addLayout(loop_layout)
        
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)
        
        # Execution controls
        execute_layout = QHBoxLayout()
        
        self.execute_btn = QPushButton("▶️ Execute Workflow")
        self.execute_btn.setStyleSheet("background-color: #4CAF50; color: white; font-size: 14px; padding: 12px;")
        self.execute_btn.clicked.connect(self.execute_workflow)
        execute_layout.addWidget(self.execute_btn)
        
        self.stop_execute_btn = QPushButton("⏹️ Stop Execution")
        self.stop_execute_btn.setStyleSheet("background-color: #f44336; color: white; font-size: 14px; padding: 12px;")
        self.stop_execute_btn.setEnabled(False)
        execute_layout.addWidget(self.stop_execute_btn)
        
        layout.addLayout(execute_layout)
        
        # Execution log
        log_group = QGroupBox("Execution Log")
        log_layout = QVBoxLayout()
        
        self.execution_log = QTextEdit()
        self.execution_log.setReadOnly(True)
        log_layout.addWidget(self.execution_log)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        return widget
        
    def create_manage_tab(self):
        """Create the management tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        list_group = QGroupBox("Saved Workflows")
        list_layout = QVBoxLayout()
        
        self.manage_workflow_list = QListWidget()
        list_layout.addWidget(self.manage_workflow_list)
        
        buttons_layout = QHBoxLayout()
        
        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.clicked.connect(self.refresh_workflow_list)
        buttons_layout.addWidget(refresh_btn)
        
        duplicate_btn = QPushButton("📋 Duplicate")
        duplicate_btn.clicked.connect(self.duplicate_workflow)
        buttons_layout.addWidget(duplicate_btn)
        
        export_btn = QPushButton("📤 Export")
        export_btn.clicked.connect(self.export_workflow)
        buttons_layout.addWidget(export_btn)
        
        import_btn = QPushButton("📥 Import")
        import_btn.clicked.connect(self.import_workflow)
        buttons_layout.addWidget(import_btn)
        
        delete_btn = QPushButton("🗑️ Delete")
        delete_btn.setStyleSheet("background-color: #f44336; color: white;")
        delete_btn.clicked.connect(self.delete_workflow)
        buttons_layout.addWidget(delete_btn)
        
        list_layout.addLayout(buttons_layout)
        list_group.setLayout(list_layout)
        layout.addWidget(list_group)
        
        info_group = QGroupBox("Workflow Information")
        info_layout = QVBoxLayout()
        
        self.workflow_info = QTextEdit()
        self.workflow_info.setReadOnly(True)
        info_layout.addWidget(self.workflow_info)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        self.manage_workflow_list.currentItemChanged.connect(self.on_manage_workflow_selected)
        
        return widget
        
    def create_advanced_tab(self):
        """Create advanced features tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Conditional actions
        conditions_group = QGroupBox("Conditional Actions")
        conditions_layout = QVBoxLayout()
        conditions_layout.addWidget(QLabel("Add if-then logic to your workflows:"))
        conditions_layout.addWidget(QLabel("• Check pixel colors"))
        conditions_layout.addWidget(QLabel("• Wait for screen changes"))
        conditions_layout.addWidget(QLabel("• Branch based on conditions"))
        add_condition_btn = QPushButton("Add Condition to Workflow")
        add_condition_btn.clicked.connect(partial(self.add_manual_action, "if_condition"))
        conditions_layout.addWidget(add_condition_btn)
        conditions_group.setLayout(conditions_layout)
        layout.addWidget(conditions_group)
        
        # Loops
        loops_group = QGroupBox("Loops & Repetition")
        loops_layout = QVBoxLayout()
        loops_layout.addWidget(QLabel("Repeat actions multiple times:"))
        loops_layout.addWidget(QLabel("• Fixed number of iterations"))
        loops_layout.addWidget(QLabel("• Nest actions within loops"))
        add_loop_btn = QPushButton("Add Loop to Workflow")
        add_loop_btn.clicked.connect(partial(self.add_manual_action, "loop"))
        loops_layout.addWidget(add_loop_btn)
        loops_group.setLayout(loops_layout)
        layout.addWidget(loops_group)
        
        # Wait actions
        wait_group = QGroupBox("Wait & Delays")
        wait_layout = QVBoxLayout()
        wait_layout.addWidget(QLabel("Add pauses in your workflow:"))
        wait_layout.addWidget(QLabel("• Fixed time delays"))
        wait_layout.addWidget(QLabel("• Wait for conditions"))
        add_wait_btn = QPushButton("Add Wait Action")
        add_wait_btn.clicked.connect(partial(self.add_manual_action, "wait"))
        wait_layout.addWidget(add_wait_btn)
        wait_group.setLayout(wait_layout)
        layout.addWidget(wait_group)
        
        # Screenshots
        screenshot_group = QGroupBox("Screenshots")
        screenshot_layout = QVBoxLayout()
        screenshot_layout.addWidget(QLabel("Capture screenshots during execution:"))
        screenshot_layout.addWidget(QLabel("• Visual verification"))
        screenshot_layout.addWidget(QLabel("• Error documentation"))
        add_screenshot_btn = QPushButton("Add Screenshot Action")
        add_screenshot_btn.clicked.connect(partial(self.add_manual_action, "screenshot"))
        screenshot_layout.addWidget(add_screenshot_btn)
        screenshot_group.setLayout(screenshot_layout)
        layout.addWidget(screenshot_group)
        
        layout.addStretch()
        
        # Tips
        tips_group = QGroupBox("💡 Tips")
        tips_layout = QVBoxLayout()
        tips = QTextEdit()
        tips.setReadOnly(True)
        tips.setMaximumHeight(150)
        tips.setHtml("""
        <ul>
            <li><b>Mouse Movement:</b> Enable in Record tab to capture smooth movements</li>
            <li><b>Screenshots:</b> Helpful for debugging and verification</li>
            <li><b>Loops:</b> Great for repetitive tasks like data entry</li>
            <li><b>Conditions:</b> Make workflows adaptive to different scenarios</li>
            <li><b>Manual Actions:</b> Add precise actions that are hard to record</li>
        </ul>
        """)
        tips_layout.addWidget(tips)
        tips_group.setLayout(tips_layout)
        layout.addWidget(tips_group)
        
        return widget
        
    def start_recording(self):
        """Start recording with options"""
        self.current_actions = []
        self.actions_list.clear()
        
        capture_movement = self.capture_movement_check.isChecked()
        capture_screenshots = self.capture_screenshots_check.isChecked()
        
        self.recorder_thread.start_recording(capture_movement, capture_screenshots)
        
        self.start_record_btn.setEnabled(False)
        self.stop_record_btn.setEnabled(True)
        self.record_status_label.setText("Status: 🔴 Recording... (Perform your actions now)")
        self.record_status_label.setStyleSheet("font-size: 13px; padding: 5px; color: red; font-weight: bold;")
        self.statusBar().showMessage("Recording in progress...")
        
    def stop_recording(self):
        """Stop recording"""
        self.recorder_thread.stop_recording()
        
        self.start_record_btn.setEnabled(True)
        self.stop_record_btn.setEnabled(False)
        self.record_status_label.setText(f"Status: ✓ Recording stopped ({len(self.current_actions)} actions captured)")
        self.record_status_label.setStyleSheet("font-size: 13px; padding: 5px; color: green;")
        self.statusBar().showMessage(f"Recording stopped. {len(self.current_actions)} actions captured")
        
    def on_action_recorded(self, action):
        """Handle recorded action"""
        self.current_actions.append(action)
        
        if action["type"] == "click":
            text = f"[{len(self.current_actions)}] Click at ({action['x']}, {action['y']}) - {action['button']}"
        elif action["type"] == "mouse_move":
            text = f"[{len(self.current_actions)}] Move to ({action['x']}, {action['y']})"
        else:
            text = f"[{len(self.current_actions)}] Key press: {action['key']}"
            
        self.actions_list.addItem(text)
        
    def add_manual_action(self, action_type="click"):
        """Open dialog to add manual action"""
        dialog = ActionDialog(self, action_type)
        if dialog.exec() == QDialog.Accepted:
            action = dialog.get_action()
            self.current_actions.append(action)
            
            # Add to list
            action_text = self.format_action_text(len(self.current_actions), action)
            self.actions_list.addItem(action_text)
            
    def format_action_text(self, index, action):
        """Format action for display"""
        action_type = action["type"]
        
        if action_type == "click":
            return f"[{index}] Click at ({action['x']}, {action['y']}) - {action.get('button', 'left')}"
        elif action_type == "mouse_move":
            return f"[{index}] Move to ({action['x']}, {action['y']})"
        elif action_type == "key":
            return f"[{index}] Key: {action['key']}"
        elif action_type == "type_text":
            return f"[{index}] Type: {action['text'][:30]}..."
        elif action_type == "wait":
            return f"[{index}] Wait: {action['duration']} seconds"
        elif action_type == "screenshot":
            return f"[{index}] Take Screenshot"
        elif action_type == "if_condition":
            return f"[{index}] IF: {action.get('condition', 'N/A')}"
        elif action_type == "loop":
            return f"[{index}] LOOP: {action.get('iterations', 1)} times"
        elif action_type == "web_navigate":
            return f"[{index}] Navigate to: {action.get('url', 'N/A')}"
        elif action_type == "web_click":
            return f"[{index}] Web Click: {action.get('selector', 'N/A')}"
        elif action_type == "web_type":
            return f"[{index}] Web Type: {action.get('text', '')[:30]} into {action.get('selector', 'N/A')}"
        elif action_type == "web_extract":
            return f"[{index}] Extract to: {action.get('variable', 'N/A')}"
        elif action_type == "excel_read":
            return f"[{index}] Read Excel: {action.get('file_path', 'N/A')}"
        elif action_type == "excel_write":
            return f"[{index}] Write Excel: {action.get('file_path', 'N/A')}"
        elif action_type == "csv_read":
            return f"[{index}] Read CSV: {action.get('file_path', 'N/A')}"
        elif action_type == "csv_write":
            return f"[{index}] Write CSV: {action.get('file_path', 'N/A')}"
        else:
            return f"[{index}] {action_type}"
            
    def edit_action(self):
        """Edit selected action"""
        current_row = self.actions_list.currentRow()
        if current_row < 0 or current_row >= len(self.current_actions):
            QMessageBox.warning(self, "Warning", "Please select an action to edit!")
            return
            
        action = self.current_actions[current_row]
        dialog = ActionDialog(self, action["type"])
        
        # Pre-fill dialog with current values
        if action["type"] == "click":
            dialog.x_spin.setValue(action.get("x", 0))
            dialog.y_spin.setValue(action.get("y", 0))
            dialog.button_combo.setCurrentText(action.get("button", "left"))
        elif action["type"] == "wait":
            dialog.wait_spin.setValue(action.get("duration", 2))
            
        if dialog.exec() == QDialog.Accepted:
            self.current_actions[current_row] = dialog.get_action()
            action_text = self.format_action_text(current_row + 1, self.current_actions[current_row])
            self.actions_list.item(current_row).setText(action_text)
            
    def delete_action(self):
        """Delete selected action"""
        current_row = self.actions_list.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "Warning", "Please select an action to delete!")
            return
            
        del self.current_actions[current_row]
        self.actions_list.takeItem(current_row)
        
        # Re-number remaining actions
        for i in range(current_row, self.actions_list.count()):
            action = self.current_actions[i]
            action_text = self.format_action_text(i + 1, action)
            self.actions_list.item(i).setText(action_text)
            
    def clear_actions(self):
        """Clear all actions"""
        self.current_actions = []
        self.actions_list.clear()
        self.record_status_label.setText("Status: Actions cleared")
        
    def save_workflow(self):
        """Save current workflow"""
        name = self.workflow_name_input.text().strip()
        
        if not name:
            QMessageBox.warning(self, "Warning", "Please enter a workflow name!")
            return
            
        if not self.current_actions:
            QMessageBox.warning(self, "Warning", "No actions to save!")
            return
            
        workflow = {
            "name": name,
            "actions": self.current_actions,
            "created": datetime.now().isoformat(),
            "action_count": len(self.current_actions)
        }
        
        self.workflows.append(workflow)
        self.save_workflows_to_file()
        
        self.workflow_name_input.clear()
        QMessageBox.information(self, "Success", f"Workflow '{name}' saved successfully!")
        
        self.refresh_workflow_list()
        self.update_workflow_combo()
        
    def load_workflows(self):
        """Load workflows from file"""
        if self.workflows_file.exists():
            try:
                with open(self.workflows_file, 'r') as f:
                    self.workflows = json.load(f)
                self.update_workflow_combo()
                self.refresh_workflow_list()
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not load workflows: {str(e)}")
                
    def save_workflows_to_file(self):
        """Save workflows to file"""
        try:
            with open(self.workflows_file, 'w') as f:
                json.dump(self.workflows, f, indent=2)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not save workflows: {str(e)}")
            
    def update_workflow_combo(self):
        """Update workflow combo box"""
        self.workflow_combo.clear()
        for workflow in self.workflows:
            self.workflow_combo.addItem(workflow["name"])
            
    def refresh_workflow_list(self):
        """Refresh the workflow list in manage tab"""
        self.manage_workflow_list.clear()
        for workflow in self.workflows:
            item_text = f"{workflow['name']} ({workflow['action_count']} actions)"
            self.manage_workflow_list.addItem(item_text)
            
    def on_workflow_selected(self, index):
        """Handle workflow selection in execute tab"""
        if index < 0 or index >= len(self.workflows):
            return
            
        workflow = self.workflows[index]
        
        details = f"Name: {workflow['name']}\n"
        details += f"Created: {workflow['created']}\n"
        details += f"Actions: {workflow['action_count']}\n\n"
        details += "Action List:\n"
        
        for i, action in enumerate(workflow['actions'], 1):
            details += f"  {self.format_action_text(i, action)}\n"
                
        self.workflow_details.setText(details)
        
    def on_manage_workflow_selected(self, current, previous):
        """Handle workflow selection in manage tab"""
        if not current:
            return
            
        index = self.manage_workflow_list.row(current)
        if index < 0 or index >= len(self.workflows):
            return
            
        workflow = self.workflows[index]
        
        info = f"Name: {workflow['name']}\n"
        info += f"Created: {workflow['created']}\n"
        info += f"Total Actions: {workflow['action_count']}\n\n"
        info += "Actions:\n"
        
        for i, action in enumerate(workflow['actions'], 1):
            info += f"  {self.format_action_text(i, action)}\n"
                
        self.workflow_info.setText(info)
        
    def delete_workflow(self):
        """Delete selected workflow"""
        current_item = self.manage_workflow_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "Warning", "Please select a workflow to delete!")
            return
            
        index = self.manage_workflow_list.row(current_item)
        workflow_name = self.workflows[index]['name']
        
        reply = QMessageBox.question(
            self, 
            "Confirm Delete",
            f"Are you sure you want to delete workflow '{workflow_name}'?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            del self.workflows[index]
            self.save_workflows_to_file()
            self.refresh_workflow_list()
            self.update_workflow_combo()
            QMessageBox.information(self, "Success", "Workflow deleted successfully!")
            
    def duplicate_workflow(self):
        """Duplicate selected workflow"""
        current_item = self.manage_workflow_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "Warning", "Please select a workflow to duplicate!")
            return
            
        index = self.manage_workflow_list.row(current_item)
        workflow = self.workflows[index].copy()
        
        workflow["name"] = workflow["name"] + " (Copy)"
        workflow["created"] = datetime.now().isoformat()
        workflow["actions"] = workflow["actions"].copy()
        
        self.workflows.append(workflow)
        self.save_workflows_to_file()
        self.refresh_workflow_list()
        self.update_workflow_combo()
        QMessageBox.information(self, "Success", "Workflow duplicated successfully!")
        
    def export_workflow(self):
        """Export workflow to JSON file"""
        current_item = self.manage_workflow_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "Warning", "Please select a workflow to export!")
            return
            
        index = self.manage_workflow_list.row(current_item)
        workflow = self.workflows[index]
        
        from PySide6.QtWidgets import QFileDialog
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Export Workflow",
            f"{workflow['name']}.json",
            "JSON Files (*.json)"
        )
        
        if filename:
            try:
                with open(filename, 'w') as f:
                    json.dump(workflow, f, indent=2)
                QMessageBox.information(self, "Success", f"Workflow exported to {filename}")
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not export workflow: {str(e)}")
                
    def import_workflow(self):
        """Import workflow from JSON file"""
        from PySide6.QtWidgets import QFileDialog
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Import Workflow",
            "",
            "JSON Files (*.json)"
        )
        
        if filename:
            try:
                with open(filename, 'r') as f:
                    workflow = json.load(f)
                    
                # Validate workflow structure
                if "name" not in workflow or "actions" not in workflow:
                    raise ValueError("Invalid workflow file")
                    
                # Update metadata
                workflow["created"] = datetime.now().isoformat()
                workflow["action_count"] = len(workflow["actions"])
                
                self.workflows.append(workflow)
                self.save_workflows_to_file()
                self.refresh_workflow_list()
                self.update_workflow_combo()
                QMessageBox.information(self, "Success", f"Workflow '{workflow['name']}' imported successfully!")
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not import workflow: {str(e)}")
            
    def execute_workflow(self):
        """Execute selected workflow"""
        index = self.workflow_combo.currentIndex()
        if index < 0:
            QMessageBox.warning(self, "Warning", "Please select a workflow to execute!")
            return
            
        workflow = self.workflows[index]
        delay = self.delay_spinbox.value()
        
        # Handle loop execution
        if self.loop_check.isChecked():
            loop_count = self.loop_count.value()
            # Wrap workflow in a loop
            workflow = {
                "name": workflow["name"],
                "actions": [{
                    "type": "loop",
                    "iterations": loop_count,
                    "actions": workflow["actions"]
                }]
            }
        
        self.execution_log.clear()
        self.execution_log.append(f"Starting execution of '{self.workflows[index]['name']}'...\n")
        self.execute_btn.setEnabled(False)
        self.stop_execute_btn.setEnabled(True)
        
        # Start executor thread
        self.executor_thread = ExecutorThread(workflow, delay)
        self.executor_thread.progress_update.connect(self.on_execution_progress)
        self.executor_thread.execution_complete.connect(self.on_execution_complete)
        self.executor_thread.screenshot_captured.connect(self.on_screenshot_captured)
        self.executor_thread.start()
        
    def on_execution_progress(self, message):
        """Handle execution progress update"""
        self.execution_log.append(message)
        self.statusBar().showMessage(message)
        
    def on_execution_complete(self, success, message):
        """Handle execution completion"""
        self.execution_log.append(f"\n{message}")
        self.execute_btn.setEnabled(True)
        self.stop_execute_btn.setEnabled(False)
        self.statusBar().showMessage(message)
        
        if success:
            QMessageBox.information(self, "Success", message)
        else:
            QMessageBox.warning(self, "Error", message)
            
    def on_screenshot_captured(self, pixmap):
        """Handle screenshot capture"""
        # Display screenshot in a new window
        screenshot_dialog = QDialog(self)
        screenshot_dialog.setWindowTitle("Screenshot Captured")
        layout = QVBoxLayout(screenshot_dialog)
        
        label = QLabel()
        scaled_pixmap = pixmap.scaled(800, 600, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        label.setPixmap(scaled_pixmap)
        layout.addWidget(label)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(screenshot_dialog.accept)
        layout.addWidget(close_btn)
        
        screenshot_dialog.exec()


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # Set a modern color palette with black text
    from PySide6.QtGui import QPalette, QColor
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(240, 240, 240))
    palette.setColor(QPalette.WindowText, QColor(0, 0, 0))  # Black text
    palette.setColor(QPalette.Base, QColor(255, 255, 255))
    palette.setColor(QPalette.AlternateBase, QColor(245, 245, 245))
    palette.setColor(QPalette.Text, QColor(0, 0, 0))  # ADD THIS - Black text in input fields
    palette.setColor(QPalette.Button, QColor(240, 240, 240))
    palette.setColor(QPalette.ButtonText, QColor(0, 0, 0))  # Black button text
    palette.setColor(QPalette.Highlight, QColor(76, 175, 80))
    palette.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
    app.setPalette(palette)
    
    window = RPAMainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()