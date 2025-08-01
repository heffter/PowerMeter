import random
import time
import csv
import os
import json
from typing import BinaryIO, Optional, List, Tuple
import threading

import pyvisa
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from flask import Flask, jsonify, request
from flask_cors import CORS


# Configuration management
def get_config_path():
    """Return the path to the configuration file."""
    # Store in the same directory as the application
    app_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(app_dir, "config.json")


# Default configuration
DEFAULT_CONFIG = {
    "device": {
        "connection_string": ""
    },
    "measurement": {
        "frequency_Hz": 1.0e9,  # 1 GHz
        "averaging": 1,
        "unit": "dBm",
        "trigger_mode": "AUTO",
        "range": "AUTO",
        "integration_time_s": 0.1,
        "channel_a": 1,
        "channel_b": 2
    },
    "display": {
        "update_frequency_Hz": 1.0,  # 1 Hz
        "time_window_s": 60,  # 60 seconds
        "auto_scale": True,
        "y_min": 0,
        "y_max": 1000
    },
    "api_server": {
        "enabled": False,
        "port": 5000,
        "host": "0.0.0.0"
    }
}


def save_config(config):
    """Save configuration to JSON file."""
    try:
        with open(get_config_path(), 'w') as f:
            json.dump(config, f, indent=4)
        return True
    except Exception as e:
        print(f"Error saving configuration: {e}")
        return False


def load_config():
    """Load configuration from JSON file or return defaults if not found."""
    config_path = get_config_path()
    
    if not os.path.exists(config_path):
        return DEFAULT_CONFIG.copy()
        
    try:
        with open(config_path, 'r') as f:
            config = json.load(f)
            
        # Ensure all required keys exist by merging with defaults
        merged_config = DEFAULT_CONFIG.copy()
        
        # Update with loaded values (only for keys that exist in DEFAULT_CONFIG)
        for section in DEFAULT_CONFIG:
            if section in config:
                for key in DEFAULT_CONFIG[section]:
                    if key in config[section]:
                        merged_config[section][key] = config[section][key]
                        
        return merged_config
    except Exception as e:
        print(f"Error loading configuration: {e}")
        return DEFAULT_CONFIG.copy()


class PowerMonitor:
    def __init__(self, root):
        self.root = root
        
        # Load configuration
        self.config = load_config()
        
        # Data structure for both channels: (timestamp, forward_power, reflected_power)
        self.data: List[Tuple[float, float, float]] = []
        self.device_connected = False
        self.simulation_mode = False
        self.n1914a = None
        self.rm = None
        self.monitoring = False
        
        # API Server state
        self.api_server = None
        self.api_thread = None
        self.api_running = False
        
        # Initialize with loaded configuration
        self.acquisition_frequency_ms = int(self.config["display"]["update_frequency_Hz"] * 1000)  # Convert Hz to ms
        
        # Initialize VISA resource manager
        try:
            self.rm = pyvisa.ResourceManager()
        except Exception as e:
            self.simulation_mode = True
        
        # Try to connect to device, fall back to simulation if failed (no popups)
        if not self.connect_to_device():
            self.simulation_mode = True
            self.device_connected = False
        else:
            # Initialize continuous measurement mode for real device
            self.initialize_continuous_measurement()
        
        self.setup_gui()
        self.setup_cleanup()
        self.start_monitoring()
        
        # Start API server if enabled in config
        if self.config["api_server"]["enabled"]:
            self.start_api_server()

    def setup_api_server(self):
        """Setup Flask API server with routes"""
        self.api_server = Flask(__name__)
        CORS(self.api_server)
        
        @self.api_server.route('/api/status', methods=['GET'])
        def get_status():
            return jsonify({
                'device_connected': self.device_connected,
                'simulation_mode': self.simulation_mode,
                'monitoring': self.monitoring,
                'acquisition_frequency_ms': self.acquisition_frequency_ms,
                'data_points': len(self.data)
            })
        
        @self.api_server.route('/api/current', methods=['GET'])
        def get_current_power():
            if not self.data:
                return jsonify({
                    'timestamp': time.time(),
                    'forward_power': 0.0,
                    'reflected_power': 0.0,
                    'vswr': 1.0
                })
            
            timestamp, forward_power, reflected_power = self.data[-1]
            
            # Calculate VSWR - handle negative reflected power by using absolute value
            reflected_power_abs = abs(reflected_power)
            if forward_power > 0 and reflected_power_abs < forward_power:
                # VSWR = (1 + sqrt(Pr/Pf)) / (1 - sqrt(Pr/Pf))
                # where Pr = reflected power, Pf = forward power
                ratio = reflected_power_abs / forward_power
                if ratio < 1:  # Valid case
                    vswr = (1 + ratio**0.5) / (1 - ratio**0.5)
                else:
                    vswr = float('inf')  # Invalid case
            else:
                vswr = 1.0  # Perfect match or no forward power
            
            return jsonify({
                'timestamp': timestamp,
                'forward_power': forward_power,
                'reflected_power': reflected_power,
                'vswr': vswr
            })
        
        @self.api_server.route('/api/history', methods=['GET'])
        def get_power_history():
            limit = request.args.get('limit', type=int, default=100)
            limit = min(limit, len(self.data))
            
            if limit <= 0:
                return jsonify([])
            
            recent_data = self.data[-limit:]
            history = []
            
            for ts, fp, rp in recent_data:
                # Calculate VSWR for each data point
                reflected_power_abs = abs(rp)
                if fp > 0 and reflected_power_abs < fp:
                    ratio = reflected_power_abs / fp
                    if ratio < 1:
                        vswr = (1 + ratio**0.5) / (1 - ratio**0.5)
                    else:
                        vswr = float('inf')
                else:
                    vswr = 1.0
                
                history.append({
                    'timestamp': ts,
                    'forward_power': fp,
                    'reflected_power': rp,
                    'vswr': vswr
                })
            
            return jsonify(history)
        
        @self.api_server.route('/api/devices', methods=['GET'])
        def list_devices():
            try:
                resources = self.rm.list_resources()
                devices = []
                
                for resource in resources:
                    try:
                        instrument = self.rm.open_resource(resource)
                        identity = instrument.query("*IDN?").strip()
                        instrument.close()
                        devices.append({
                            'resource': resource,
                            'identity': identity,
                            'is_n1914a': 'N1914A' in identity
                        })
                    except:
                        devices.append({
                            'resource': resource,
                            'identity': 'Unknown/Error',
                            'is_n1914a': False
                        })
                
                return jsonify({
                    'success': True,
                    'devices': devices
                })
                
            except Exception as e:
                return jsonify({
                    'success': False,
                    'message': f'Failed to list devices: {str(e)}'
                }), 500

    def start_api_server(self):
        """Start the API server in a separate thread"""
        if self.api_running:
            return
        
        try:
            self.setup_api_server()
            host = self.config["api_server"]["host"]
            port = self.config["api_server"]["port"]
            
            def run_server():
                self.api_server.run(host=host, port=port, debug=False, use_reloader=False)
            
            self.api_thread = threading.Thread(target=run_server, daemon=True)
            self.api_thread.start()
            self.api_running = True
            
            print(f"API Server started on {host}:{port}")
            
        except Exception as e:
            print(f"Failed to start API server: {e}")
            messagebox.showerror("API Server Error", f"Failed to start API server: {str(e)}")

    def stop_api_server(self):
        """Stop the API server"""
        if not self.api_running:
            return
        
        try:
            # Flask doesn't have a clean shutdown method, so we'll just mark it as stopped
            self.api_running = False
            self.api_server = None
            self.api_thread = None
            print("API Server stopped")
            
        except Exception as e:
            print(f"Error stopping API server: {e}")

    def toggle_api_server(self):
        """Toggle the API server on/off"""
        if self.api_running:
            self.stop_api_server()
            self.config["api_server"]["enabled"] = False
            self.api_toggle_btn.config(text="Start API Server")
            messagebox.showinfo("API Server", "API Server stopped")
        else:
            self.start_api_server()
            if self.api_running:
                self.config["api_server"]["enabled"] = True
                self.api_toggle_btn.config(text="Stop API Server")
                messagebox.showinfo("API Server", f"API Server started on port {self.config['api_server']['port']}")
        
        save_config(self.config)

    def initialize_continuous_measurement(self):
        """Initialize continuous measurement mode for both channels"""
        if not self.n1914a or not self.device_connected:
            return
        
        try:
            # Set continuous measurement mode for both channels (only once during initialization)
            self.n1914a.write(":INITiate1:CONTinuous ON")
            self.n1914a.write(":INITiate2:CONTinuous ON")
            print("Continuous measurement mode initialized for both channels")
        except Exception as e:
            print(f"Error initializing continuous measurement: {e}")

    def connect_to_device(self) -> bool:
        # Use connection string from config
        connection_string = self.config["device"]["connection_string"]
        
        # If connection string is empty, try to detect device
        if not connection_string:
            try:
                resources = self.rm.list_resources()
                if not resources:
                    return False
                for resource in resources:
                    try:
                        instrument = self.rm.open_resource(resource)
                        identity = instrument.query("*IDN?").strip()
                        if "N1914A" in identity:
                            self.n1914a = instrument
                            self.device_connected = True
                            # Save successful connection string
                            self.config["device"]["connection_string"] = resource
                            save_config(self.config)
                            return True
                        instrument.close()
                    except:
                        continue
                return False
            except Exception as e:
                return False
        else:
            # Try to connect using saved connection string
            try:
                self.n1914a = self.rm.open_resource(connection_string)
                identity = self.n1914a.query("*IDN?").strip()
                if "N1914A" in identity:
                    self.device_connected = True
                    return True
                else:
                    self.n1914a.close()
                    self.n1914a = None
                    return False
            except Exception as e:
                return False

    def setup_gui(self):
        self.root.title("Power Monitor - Keysight N1914A")
        self.root.geometry("1000x800")
        self.root.configure(bg='#f5f7fa')

        # Main container
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Status and control frame
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(0, 10))

        # Control buttons (top right)
        control_right = ttk.Frame(status_frame)
        control_right.pack(side=tk.RIGHT)
        self.toggle_btn = ttk.Button(control_right, text="Connect to Device",
                                    command=self.toggle_simulation_mode)
        self.toggle_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # API Server toggle button
        api_button_text = "Stop API Server" if self.config["api_server"]["enabled"] else "Start API Server"
        self.api_toggle_btn = ttk.Button(control_right, text=api_button_text,
                                        command=self.toggle_api_server)
        self.api_toggle_btn.pack(side=tk.LEFT, padx=(0, 10))

        # Acquisition frequency control
        freq_frame = ttk.Frame(main_frame)
        freq_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(freq_frame, text="Acquisition Frequency (ms):").pack(side=tk.LEFT, padx=(0, 10))
        self.freq_var = tk.StringVar(value=str(self.acquisition_frequency_ms))
        self.freq_spinbox = ttk.Spinbox(freq_frame, from_=100, to=10000, increment=100,
                                       textvariable=self.freq_var, width=10,
                                       command=self.update_acquisition_frequency)
        self.freq_spinbox.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(freq_frame, text="Apply", command=self.update_acquisition_frequency).pack(side=tk.LEFT, padx=(0, 20))
        
        # Y-axis scaling control
        ttk.Label(freq_frame, text="Y-Axis Auto-Scale:").pack(side=tk.LEFT, padx=(0, 5))
        self.auto_scale_var = tk.BooleanVar(value=self.config["display"]["auto_scale"])
        self.auto_scale_check = ttk.Checkbutton(freq_frame, text="Enabled", variable=self.auto_scale_var,
                                               command=self.toggle_auto_scale)
        self.auto_scale_check.pack(side=tk.LEFT)
        
        # Manual Y-axis range controls (initially hidden)
        self.manual_scale_frame = ttk.Frame(freq_frame)
        ttk.Label(self.manual_scale_frame, text="Y-Min:").pack(side=tk.LEFT, padx=(0, 2))
        self.y_min_var = tk.StringVar(value=str(self.config["display"]["y_min"]))
        self.y_min_entry = ttk.Entry(self.manual_scale_frame, textvariable=self.y_min_var, width=8)
        self.y_min_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(self.manual_scale_frame, text="Y-Max:").pack(side=tk.LEFT, padx=(0, 2))
        self.y_max_var = tk.StringVar(value=str(self.config["display"]["y_max"]))
        self.y_max_entry = ttk.Entry(self.manual_scale_frame, textvariable=self.y_max_var, width=8)
        self.y_max_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(self.manual_scale_frame, text="Apply Range", command=self.apply_manual_scale).pack(side=tk.LEFT)
        
        # Show manual controls if auto-scale is disabled
        if not self.auto_scale_var.get():
            self.manual_scale_frame.pack(side=tk.LEFT, padx=(10, 0))

        # Current power display
        current_frame = ttk.Frame(main_frame, style='Card.TFrame')
        current_frame.pack(fill=tk.X, pady=10, ipadx=10, ipady=10)
        
        # Forward Power (Channel A)
        forward_frame = ttk.Frame(current_frame)
        forward_frame.pack(side=tk.LEFT, padx=(0, 20))
        ttk.Label(forward_frame, text="FORWARD POWER:",
                  font=('Helvetica', 10), foreground='#6c757d').pack(side=tk.LEFT)
        self.forward_power_var = tk.StringVar(value="0.00 W")
        ttk.Label(forward_frame, textvariable=self.forward_power_var,
                  font=('Helvetica', 28, 'bold'), foreground='#2c7be5').pack(side=tk.LEFT, padx=10)
        
        # Reflected Power (Channel B)
        reflected_frame = ttk.Frame(current_frame)
        reflected_frame.pack(side=tk.LEFT)
        ttk.Label(reflected_frame, text="REFLECTED POWER:",
                  font=('Helvetica', 10), foreground='#6c757d').pack(side=tk.LEFT)
        self.reflected_power_var = tk.StringVar(value="0.00 W")
        ttk.Label(reflected_frame, textvariable=self.reflected_power_var,
                  font=('Helvetica', 28, 'bold'), foreground='#dc3545').pack(side=tk.LEFT, padx=10)

        # Graph frame
        graph_frame = ttk.Frame(main_frame, style='Card.TFrame')
        graph_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.figure = plt.Figure(figsize=(8, 5), dpi=100, facecolor='none')
        self.ax = self.figure.add_subplot(111)
        plt.style.use('seaborn-v0_8-whitegrid')
        self.ax.set_facecolor('#ffffff')
        self.figure.patch.set_alpha(0)
        self.canvas = FigureCanvasTkAgg(self.figure, master=graph_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Controls
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X)
        ttk.Button(control_frame, text="Export Data",
                   style='Accent.TButton', command=self.export_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Configure Device",
                   command=self.configure_device).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Exit",
                   command=self.cleanup_and_exit).pack(side=tk.RIGHT, padx=5)

        # Bottom status bar for connection status
        self.statusbar = ttk.Frame(self.root)
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)
        # Place connection_status_label to the left of connection_status, both at bottom right
        self.connection_status = ttk.Label(self.statusbar, font=('Helvetica', 10, 'bold'))
        self.connection_status.pack(side=tk.RIGHT, padx=(0, 10))
        self.connection_status_label = ttk.Label(self.statusbar, text="Connection Status:", font=('Helvetica', 10, 'bold'), foreground='black')
        self.connection_status_label.pack(side=tk.RIGHT, padx=(0, 2))
        self.update_status_display()

    def update_acquisition_frequency(self):
        try:
            new_freq = int(self.freq_var.get())
            if 100 <= new_freq <= 10000:
                self.acquisition_frequency_ms = new_freq
                # Update and save configuration
                self.config["display"]["update_frequency_Hz"] = new_freq / 1000.0  # Convert ms to Hz
                save_config(self.config)
                messagebox.showinfo("Success", f"Acquisition frequency updated to {new_freq}ms")
            else:
                messagebox.showerror("Error", "Frequency must be between 100ms and 10000ms")
                self.freq_var.set(str(self.acquisition_frequency_ms))
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number")
            self.freq_var.set(str(self.acquisition_frequency_ms))

    def toggle_auto_scale(self):
        """Toggle between auto-scaling and manual Y-axis control"""
        if self.auto_scale_var.get():
            # Auto-scale enabled - hide manual controls
            self.manual_scale_frame.pack_forget()
        else:
            # Auto-scale disabled - show manual controls
            self.manual_scale_frame.pack(side=tk.LEFT, padx=(10, 0))
        
        # Save setting to configuration
        self.config["display"]["auto_scale"] = self.auto_scale_var.get()
        save_config(self.config)
        
        self.update_gui()

    def apply_manual_scale(self):
        """Apply manual Y-axis range settings"""
        try:
            y_min = float(self.y_min_var.get())
            y_max = float(self.y_max_var.get())
            
            if y_min >= y_max:
                messagebox.showerror("Error", "Y-Min must be less than Y-Max")
                return
            
            if y_min < 0:
                messagebox.showerror("Error", "Y-Min cannot be negative")
                return
            
            # Apply the manual range
            self.ax.set_ylim(y_min, y_max)
            self.canvas.draw()
            
            # Save settings to configuration
            self.config["display"]["y_min"] = y_min
            self.config["display"]["y_max"] = y_max
            save_config(self.config)
            
            messagebox.showinfo("Success", f"Y-axis range set to {y_min:.1f} - {y_max:.1f}")
            
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numbers for Y-axis range")

    def setup_cleanup(self):
        self.root.protocol("WM_DELETE_WINDOW", self.cleanup_and_exit)

    def toggle_simulation_mode(self):
        if self.simulation_mode:
            if self.connect_to_device():
                self.simulation_mode = False
                self.toggle_btn.config(text="Disconnect Device")
                # Clear data when switching to real mode
                self.data.clear()
                messagebox.showinfo("Success", "Connected to real device!")
            else:
                messagebox.showerror("Error", "No device found. Staying in simulation mode.")
        else:
            self.simulation_mode = True
            self.device_connected = False
            if self.n1914a:
                self.n1914a.close()
                self.n1914a = None
            self.toggle_btn.config(text="Connect to Device")
            # Clear data when switching to simulation mode
            self.data.clear()
            messagebox.showinfo("Info", "Switched to simulation mode.")
        self.update_status_display()

    def update_status_display(self):
        if self.device_connected and not self.simulation_mode:
            self.connection_status.config(text="Connected", foreground='#28a745')
        else:
            self.connection_status.config(text="Simulation mode", foreground='#dc3545')

    def configure_device(self):
        config_window = tk.Toplevel(self.root)
        config_window.title("Device Configuration")
        config_window.geometry("600x700")
        config_window.transient(self.root)
        config_window.grab_set()
        
        # Main scrollable frame
        main_canvas = tk.Canvas(config_window)
        scrollbar = ttk.Scrollbar(config_window, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        
        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Title
        ttk.Label(scrollable_frame, text="Device Configuration", font=('Helvetica', 16, 'bold')).pack(pady=10)
        
        # Device Connection Section
        connection_frame = ttk.LabelFrame(scrollable_frame, text="Device Connection", padding=10)
        connection_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # USB Device Selection
        ttk.Label(connection_frame, text="USB Device Selection:", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W, pady=(0, 5))
        
        # Available devices listbox
        devices_frame = ttk.Frame(connection_frame)
        devices_frame.pack(fill=tk.X, pady=5)
        ttk.Label(devices_frame, text="Available Devices:").pack(anchor=tk.W)
        
        device_listbox = tk.Listbox(devices_frame, height=4, exportselection=0)
        device_listbox.pack(fill=tk.X, pady=(0, 5))
        
        # Scan and refresh buttons
        scan_frame = ttk.Frame(connection_frame)
        scan_frame.pack(fill=tk.X, pady=5)
        
        def scan_devices():
            device_listbox.delete(0, tk.END)
            try:
                resources = self.rm.list_resources()
                for i, resource in enumerate(resources):
                    try:
                        instrument = self.rm.open_resource(resource)
                        identity = instrument.query("*IDN?").strip()
                        instrument.close()
                    except:
                        identity = "Unknown/Error"
                    device_listbox.insert(tk.END, f"{resource}  |  ID: {identity}")
            except Exception as e:
                device_listbox.insert(tk.END, f"Error scanning devices: {str(e)}")
        
        ttk.Button(scan_frame, text="Scan for Devices", command=scan_devices).pack(side=tk.LEFT, padx=(0, 10))
        
        # Manual connection entry
        manual_frame = ttk.Frame(connection_frame)
        manual_frame.pack(fill=tk.X, pady=5)
        ttk.Label(manual_frame, text="Manual Connection String:").pack(anchor=tk.W)
        # Use the last connection string from config, or default if not available
        last_connection = self.config["device"]["connection_string"] if self.config["device"]["connection_string"] else "USB0::0x0957::0x0607::MY12345678::INSTR"
        manual_var = tk.StringVar(value=last_connection)
        manual_entry = ttk.Entry(manual_frame, textvariable=manual_var, width=50)
        manual_entry.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(manual_frame, text="Format: USB0::ManufacturerID::ModelID::Serial::INSTR\nExample: USB0::0x0957::0x0607::MY12345678::INSTR", font=('Helvetica', 8), foreground='gray').pack(anchor=tk.W)
        
        # Connection status
        status_var = tk.StringVar(value="Not Connected")
        status_label = ttk.Label(connection_frame, textvariable=status_var, font=('Helvetica', 10, 'bold'))
        status_label.pack(pady=5)
        
        # Test Connection button
        def test_connection():
            try:
                if device_listbox.curselection():
                    # Use selected device
                    selection = device_listbox.curselection()[0]
                    resources = self.rm.list_resources()
                    if selection < len(resources):
                        resource = resources[selection]
                    else:
                        status_var.set("Error: Invalid selection")
                        return
                else:
                    # Use manual entry
                    resource = manual_var.get()
                
                instrument = self.rm.open_resource(resource)
                identity = instrument.query("*IDN?").strip()
                instrument.close()
                status_var.set(f"Connected: {identity}")
            except Exception as e:
                status_var.set(f"Connection Failed: {str(e)}")
        
        ttk.Button(connection_frame, text="Test Connection", command=test_connection).pack(pady=5)
        
        # Device Configuration Section
        config_frame = ttk.LabelFrame(scrollable_frame, text="Device Configuration", padding=10)
        config_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Frequency Configuration
        freq_frame = ttk.Frame(config_frame)
        freq_frame.pack(fill=tk.X, pady=5)
        ttk.Label(freq_frame, text="Frequency (Hz):", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W)
        freq_var = tk.StringVar(value=str(int(self.config["measurement"]["frequency_Hz"])))
        freq_entry = ttk.Entry(freq_frame, textvariable=freq_var, width=20)
        freq_entry.pack(anchor=tk.W, pady=(0, 5))
        ttk.Label(freq_frame, text="Range: 1 Hz to 50 GHz", font=('Helvetica', 8), foreground='gray').pack(anchor=tk.W)
        
        # Averaging Configuration
        avg_frame = ttk.Frame(config_frame)
        avg_frame.pack(fill=tk.X, pady=5)
        ttk.Label(avg_frame, text="Averaging Count:", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W)
        avg_var = tk.StringVar(value=str(self.config["measurement"]["averaging"]))
        avg_entry = ttk.Entry(avg_frame, textvariable=avg_var, width=20)
        avg_entry.pack(anchor=tk.W, pady=(0, 5))
        ttk.Label(avg_frame, text="Range: 1 to 1000", font=('Helvetica', 8), foreground='gray').pack(anchor=tk.W)
        
        # Measurement Unit
        unit_frame = ttk.Frame(config_frame)
        unit_frame.pack(fill=tk.X, pady=5)
        ttk.Label(unit_frame, text="Measurement Unit:", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W)
        unit_var = tk.StringVar(value=self.config["measurement"]["unit"])
        unit_combo = ttk.Combobox(unit_frame, textvariable=unit_var, values=["W", "dBm", "dBW"], state="readonly", width=10)
        unit_combo.pack(anchor=tk.W, pady=(0, 5))
        
        # Trigger Configuration
        trigger_frame = ttk.Frame(config_frame)
        trigger_frame.pack(fill=tk.X, pady=5)
        ttk.Label(trigger_frame, text="Trigger Source:", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W)
        trigger_var = tk.StringVar(value=self.config["measurement"]["trigger_mode"])
        trigger_combo = ttk.Combobox(trigger_frame, textvariable=trigger_var, values=["IMM", "EXT", "BUS"], state="readonly", width=10)
        trigger_combo.pack(anchor=tk.W, pady=(0, 5))
        
        # Auto Range
        autorange_frame = ttk.Frame(config_frame)
        autorange_frame.pack(fill=tk.X, pady=5)
        autorange_var = tk.BooleanVar(value=self.config["measurement"]["range"] == "AUTO")
        ttk.Checkbutton(autorange_frame, text="Auto Range", variable=autorange_var).pack(anchor=tk.W)
        
        # Range (if auto range is off)
        range_frame = ttk.Frame(config_frame)
        range_frame.pack(fill=tk.X, pady=5)
        ttk.Label(range_frame, text="Manual Range (W):", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W)
        range_var = tk.StringVar(value=str(self.config["measurement"]["range"]) if self.config["measurement"]["range"] != "AUTO" else "1")
        range_entry = ttk.Entry(range_frame, textvariable=range_var, width=20)
        range_entry.pack(anchor=tk.W, pady=(0, 5))
        ttk.Label(range_frame, text="Range: 0.1 to 100 W", font=('Helvetica', 8), foreground='gray').pack(anchor=tk.W)
        
        # Integration Time
        integration_frame = ttk.Frame(config_frame)
        integration_frame.pack(fill=tk.X, pady=5)
        ttk.Label(integration_frame, text="Integration Time (s):", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W)
        integration_var = tk.StringVar(value=str(self.config["measurement"]["integration_time_s"]))
        integration_entry = ttk.Entry(integration_frame, textvariable=integration_var, width=20)
        integration_entry.pack(anchor=tk.W, pady=(0, 5))
        ttk.Label(integration_frame, text="Range: 0.001 to 1.0", font=('Helvetica', 8), foreground='gray').pack(anchor=tk.W)
        
        # API Server Configuration Section
        api_frame = ttk.LabelFrame(scrollable_frame, text="API Server Configuration", padding=10)
        api_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # API Server Enable/Disable
        api_enable_frame = ttk.Frame(api_frame)
        api_enable_frame.pack(fill=tk.X, pady=5)
        api_enable_var = tk.BooleanVar(value=self.config["api_server"]["enabled"])
        ttk.Checkbutton(api_enable_frame, text="Enable REST API Server", variable=api_enable_var).pack(anchor=tk.W)
        
        # API Server Port
        api_port_frame = ttk.Frame(api_frame)
        api_port_frame.pack(fill=tk.X, pady=5)
        ttk.Label(api_port_frame, text="API Server Port:", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W)
        api_port_var = tk.StringVar(value=str(self.config["api_server"]["port"]))
        api_port_entry = ttk.Entry(api_port_frame, textvariable=api_port_var, width=20)
        api_port_entry.pack(anchor=tk.W, pady=(0, 5))
        ttk.Label(api_port_frame, text="Range: 1024 to 65535", font=('Helvetica', 8), foreground='gray').pack(anchor=tk.W)
        
        # API Server Host
        api_host_frame = ttk.Frame(api_frame)
        api_host_frame.pack(fill=tk.X, pady=5)
        ttk.Label(api_host_frame, text="API Server Host:", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W)
        api_host_var = tk.StringVar(value=self.config["api_server"]["host"])
        api_host_entry = ttk.Entry(api_host_frame, textvariable=api_host_var, width=20)
        api_host_entry.pack(anchor=tk.W, pady=(0, 5))
        ttk.Label(api_host_frame, text="Use '0.0.0.0' for all interfaces, '127.0.0.1' for localhost only", font=('Helvetica', 8), foreground='gray').pack(anchor=tk.W)
        
        # Apply Configuration Section
        apply_frame = ttk.LabelFrame(scrollable_frame, text="Apply Configuration", padding=10)
        apply_frame.pack(fill=tk.X, padx=10, pady=5)
        
        def apply_config():
            try:
                # Get selected device or manual entry
                if device_listbox.curselection():
                    selection = device_listbox.curselection()[0]
                    resources = self.rm.list_resources()
                    if selection < len(resources):
                        resource = resources[selection]
                    else:
                        messagebox.showerror("Error", "Invalid device selection")
                        return
                else:
                    resource = manual_var.get()
                
                # Connect to device
                if self.n1914a:
                    self.n1914a.close()
                
                self.n1914a = self.rm.open_resource(resource)
                
                # Apply all configurations
                freq = float(freq_var.get())
                avg_count = int(avg_var.get())
                unit = unit_var.get()
                trigger = trigger_var.get()
                autorange = autorange_var.get()
                range_val = float(range_var.get())
                integration = float(integration_var.get())
                
                # Update configuration
                self.config["measurement"]["frequency_Hz"] = freq
                self.config["measurement"]["averaging"] = avg_count
                self.config["measurement"]["unit"] = unit
                self.config["measurement"]["trigger_mode"] = trigger
                self.config["measurement"]["range"] = "AUTO" if autorange else str(range_val)
                self.config["measurement"]["integration_time_s"] = integration
                
                # Set frequency
                self.n1914a.write(f"SENS:FREQ {freq}")
                
                # Configure Channel 1 (Forward Power) - using correct SCPI commands from programming guide
                self.n1914a.write(f"SENS1:AVER:COUN {avg_count}")
                self.n1914a.write(f"SENS1:UNIT:POW {unit}")
                self.n1914a.write(f"SENS1:TRIG:SOUR {trigger}")
                if autorange:
                    self.n1914a.write("SENS1:POW:RANG:AUTO ON")
                else:
                    self.n1914a.write("SENS1:POW:RANG:AUTO OFF")
                    self.n1914a.write(f"SENS1:POW:RANG {range_val}")
                self.n1914a.write(f"SENS1:POW:INT {integration}")
                
                # Configure Channel 2 (Reflected Power) - using correct SCPI commands from programming guide
                self.n1914a.write(f"SENS2:AVER:COUN {avg_count}")
                self.n1914a.write(f"SENS2:UNIT:POW {unit}")
                self.n1914a.write(f"SENS2:TRIG:SOUR {trigger}")
                if autorange:
                    self.n1914a.write("SENS2:POW:RANG:AUTO ON")
                else:
                    self.n1914a.write("SENS2:POW:RANG:AUTO OFF")
                    self.n1914a.write(f"SENS2:POW:RANG {range_val}")
                self.n1914a.write(f"SENS2:POW:INT {integration}")
                
                # Test the configuration
                identity = self.n1914a.query("*IDN?").strip()
                
                # Save successful connection string
                self.config["device"]["connection_string"] = resource
                
                # Handle API Server configuration
                api_enabled = api_enable_var.get()
                api_port = int(api_port_var.get())
                api_host = api_host_var.get()
                
                # Validate API server settings
                if api_enabled:
                    if not (1024 <= api_port <= 65535):
                        messagebox.showerror("Error", "API Server port must be between 1024 and 65535")
                        return
                    if not api_host:
                        messagebox.showerror("Error", "API Server host cannot be empty")
                        return
                
                # Update API server configuration
                self.config["api_server"]["enabled"] = api_enabled
                self.config["api_server"]["port"] = api_port
                self.config["api_server"]["host"] = api_host
                
                # Restart API server if configuration changed
                if self.api_running:
                    self.stop_api_server()
                    if api_enabled:
                        self.start_api_server()
                        if self.api_running:
                            self.api_toggle_btn.config(text="Stop API Server")
                        else:
                            self.config["api_server"]["enabled"] = False
                            api_enable_var.set(False)
                elif api_enabled:
                    self.start_api_server()
                    if self.api_running:
                        self.api_toggle_btn.config(text="Stop API Server")
                    else:
                        self.config["api_server"]["enabled"] = False
                        api_enable_var.set(False)
                
                # Save configuration
                save_config(self.config)
                
                # Update application state
                self.device_connected = True
                self.simulation_mode = False
                self.toggle_btn.config(text="Disconnect Device")
                self.data.clear()  # Clear data when switching to real device
                self.update_status_display()
                
                # Initialize continuous measurement mode
                self.initialize_continuous_measurement()
                
                messagebox.showinfo("Success", f"Device configured successfully!\nDevice: {identity}")
                status_var.set(f"Connected: {identity}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to apply configuration: {str(e)}")
                status_var.set(f"Configuration Failed: {str(e)}")
        
        ttk.Button(apply_frame, text="Apply Configuration", command=apply_config).pack(pady=10)
        
        # Note: Device scan is only performed when user clicks "Scan for Devices" button

    def generate_power_reading(self) -> Tuple[float, float]:
        # Generate forward power (Channel A)
        forward_base = 800 + random.uniform(-100, 100)
        forward_drift = random.uniform(-10, 10)
        forward_spike = random.random() > 0.95
        forward_noise = random.uniform(-20, 20)
        forward_result = forward_base + forward_drift + forward_noise
        if forward_spike:
            forward_result *= 1.5
        
        # Generate reflected power (Channel B) - typically lower than forward
        reflected_base = 50 + random.uniform(-20, 20)
        reflected_drift = random.uniform(-5, 5)
        reflected_spike = random.random() > 0.98
        reflected_noise = random.uniform(-10, 10)
        reflected_result = reflected_base + reflected_drift + reflected_noise
        if reflected_spike:
            reflected_result *= 2.0
        
        return (round(max(0, forward_result), 2), round(max(0, reflected_result), 2))

    def read_n1914a_power(self) -> Optional[Tuple[float, float]]:
        if not self.n1914a or not self.device_connected:
            return None
        try:
            # Read from Channel 1 (Forward Power) - using correct SCPI commands from programming guide
            forward_power = self.n1914a.query_ascii_values(":FETCh1:SCALar:POWer:AC?")[0]
            
            # Read from Channel 2 (Reflected Power) - using correct SCPI commands from programming guide
            reflected_power = self.n1914a.query_ascii_values(":FETCh2:SCALar:POWer:AC?")[0]
            
            return (float(forward_power), float(reflected_power))
        except Exception as e:
            print(f"Error reading power: {e}")
            self.device_connected = False
            return None

    def start_monitoring(self):
        self.monitoring = True
        self.update_data()

    def update_data(self):
        if not self.monitoring:
            return
        timestamp = time.time()
        if self.simulation_mode:
            power = self.generate_power_reading()
        else:
            power = self.read_n1914a_power()
            if power is None:
                self.simulation_mode = True
                power = self.generate_power_reading()
                self.update_status_display()
        self.data.append((timestamp, power[0], power[1])) # Append forward and reflected power
        cutoff_time = timestamp - 60.0
        self.data = [(t, f, r) for t, f, r in self.data if t >= cutoff_time]
        self.update_gui()
        self.root.after(self.acquisition_frequency_ms, self.update_data)

    def update_gui(self):
        if not self.data:
            return
        _, current_forward, current_reflected = self.data[-1]
        
        # Format power values - show 0.00 instead of -0.00 for very small negative values
        forward_display = f"{current_forward:.2f} W"
        reflected_display = f"{current_reflected:.2f} W" if abs(current_reflected) >= 0.01 else "0.00 W"
        
        self.forward_power_var.set(forward_display)
        self.reflected_power_var.set(reflected_display)
        self.ax.clear()
        timestamps = [datetime.fromtimestamp(ts) for ts, _, _ in self.data]
        forward_powers = [f for _, f, _ in self.data]
        reflected_powers = [r for _, _, r in self.data]
        
        # Plot Forward Power
        forward_line = self.ax.plot(timestamps, forward_powers,
                                    color='#2c7be5',
                                    linewidth=2.5,
                                    alpha=0.8,
                                    marker='o',
                                    markersize=5,
                                    markerfacecolor='#ffffff',
                                    markeredgecolor='#2c7be5',
                                    markeredgewidth=1.5,
                                    zorder=3,
                                    label='Forward Power')[0]
        self.ax.fill_between(timestamps, forward_powers, color='#2c7be5', alpha=0.1)
        
        # Plot Reflected Power
        reflected_line = self.ax.plot(timestamps, reflected_powers,
                                     color='#dc3545',
                                     linewidth=2.5,
                                     alpha=0.8,
                                     marker='s',
                                     markersize=5,
                                     markerfacecolor='#ffffff',
                                     markeredgecolor='#dc3545',
                                     markeredgewidth=1.5,
                                     zorder=3,
                                     label='Reflected Power')[0]
        self.ax.fill_between(timestamps, reflected_powers, color='#dc3545', alpha=0.1)

        self.ax.set_title('Real-Time Power Measurement - Keysight N1914A',
                          fontsize=12,
                          fontweight='bold',
                          pad=15,
                          color='#343a40')
        self.ax.set_ylabel('Watts (W)',
                           fontsize=10,
                           labelpad=10,
                           color='#6c757d')
        self.ax.set_xticks([])
        self.ax.set_xlabel('')
        self.ax.grid(True, linestyle=':', alpha=0.6, color='#e9ecef')
        for spine in ['top', 'right', 'left', 'bottom']:
            self.ax.spines[spine].set_color('#e9ecef')
        
        # Add legend
        self.ax.legend(loc='upper right', framealpha=0.9, fancybox=True, shadow=True)
        
        # Improved Y-axis autoscaling
        if self.auto_scale_var.get():  # Only auto-scale if enabled
            if len(forward_powers) > 0 or len(reflected_powers) > 0:
                # Get all power values
                all_powers = forward_powers + reflected_powers
                
                if all_powers:
                    min_power = min(all_powers)
                    max_power = max(all_powers)
                    power_range = max_power - min_power
                    
                    # Calculate intelligent padding based on power range
                    if power_range > 0:
                        # Use percentage-based padding for larger ranges
                        padding_factor = 0.15  # 15% padding
                        y_padding = max(power_range * padding_factor, 10)  # Minimum 10W padding
                    else:
                        # If all values are the same, add fixed padding
                        y_padding = max(max_power * 0.1, 10)  # 10% of value or 10W minimum
                    
                    # Set Y-axis limits with intelligent bounds
                    y_min = max(0, min_power - y_padding)
                    y_max = max_power + y_padding
                    
                    # Ensure we have a reasonable range even for very small values
                    if y_max - y_min < 20:
                        y_max = y_min + 20
                    
                    self.ax.set_ylim(y_min, y_max)
                    
                    # Add Y-axis ticks with appropriate spacing
                    if y_max - y_min > 100:
                        # For large ranges, use fewer ticks
                        self.ax.yaxis.set_major_locator(plt.MaxNLocator(6))
                    else:
                        # For smaller ranges, use more ticks
                        self.ax.yaxis.set_major_locator(plt.MaxNLocator(8))
            else:
                # Default range when no data is available
                self.ax.set_ylim(0, 1000)
        else:
            # Manual scaling mode - preserve current range or use default
            try:
                y_min = float(self.y_min_var.get())
                y_max = float(self.y_max_var.get())
                if y_min < y_max and y_min >= 0:
                    self.ax.set_ylim(y_min, y_max)
                else:
                    # Fallback to default range if manual values are invalid
                    self.ax.set_ylim(0, 1000)
            except ValueError:
                # Fallback to default range if manual values are invalid
                self.ax.set_ylim(0, 1000)
        
        self.figure.tight_layout()
        self.canvas.draw()

    def export_csv(self):
        if not self.data:
            messagebox.showwarning("Warning", "No data to export")
            return
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"power_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if not filename:
            return
        try:
            with open(filename, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Timestamp', 'Epoch Time', 'Forward Power (W)', 'Reflected Power (W)', 'Mode'])
                for ts, f, r in self.data:
                    mode = "Simulation" if self.simulation_mode else "Real Device"
                    writer.writerow([
                        datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S'),
                        f"{ts:.3f}",
                        f,
                        r,
                        mode
                    ])
            messagebox.showinfo("Success", f"Data exported to {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export: {str(e)}")

    def cleanup_and_exit(self):
        self.monitoring = False
        self.stop_api_server()  # Stop API server
        if self.n1914a:
            try:
                self.n1914a.close()
            except:
                pass
        if self.rm:
            try:
                self.rm.close()
            except:
                pass
        self.root.quit()


if __name__ == "__main__":
    root = tk.Tk()
    app = PowerMonitor(root)
    root.mainloop()