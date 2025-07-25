import random
import time
import csv
from typing import BinaryIO

import pyvisa
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class PowerMonitor:
    def __init__(self, root):
        self.root = root
        self.data = []
        visa_addr = "USB0::0x0957::0x1A07::MY12345678::0::INSTR"  # (example VISA address for N1914A)
        if not self.setup_n1914a(visa_addr, channel=1, frequency_hz=1e9):
            messagebox.showerror("Error", "Unable to setup Keysight N1914A Power Meter via USB")
            exit(1)
        self.setup_gui()
        self.update_data()


    def setup_gui(self):
        self.root.title("Power Monitor")
        self.root.geometry("800x620")
        self.root.configure(bg='#f5f7fa')

        # Main container
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Current power display
        current_frame = ttk.Frame(main_frame, style='Card.TFrame')
        current_frame.pack(fill=tk.X, pady=10, ipadx=10, ipady=10)

        ttk.Label(current_frame, text="CURRENT POWER:",
                  font=('Helvetica', 10), foreground='#6c757d').pack(side=tk.LEFT)
        self.current_power_var = tk.StringVar(value="0.00 W")
        ttk.Label(current_frame, textvariable=self.current_power_var,
                  font=('Helvetica', 28, 'bold'), foreground='#2c7be5').pack(side=tk.LEFT, padx=10)

        # Graph frame
        graph_frame = ttk.Frame(main_frame, style='Card.TFrame')
        graph_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Configure matplotlib figure
        self.figure = plt.Figure(figsize=(7, 4.5), dpi=100, facecolor='none')
        self.ax = self.figure.add_subplot(111)

        # Apply modern style
        plt.style.use('seaborn-v0_8-whitegrid')
        self.ax.set_facecolor('#ffffff')
        self.figure.patch.set_alpha(0)

        # Create canvas
        self.canvas = FigureCanvasTkAgg(self.figure, master=graph_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Controls
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X)

        ttk.Button(control_frame, text="Export Data",
                   style='Accent.TButton', command=self.export_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Exit",
                   command=self.root.quit).pack(side=tk.RIGHT, padx=5)

    def generate_power_reading(self):
        # Simulate more realistic power fluctuations
        base = 800 + random.uniform(-100, 100)
        spike = random.random() > 0.95  # 5% chance of spike
        return round(base * (1.5 if spike else 1), 2)

    def setup_n1914a(self, visa_address: str, channel: int = 1, frequency_hz: float = 1e9) -> bool:
        # Connect to Keysight N1914A

        # Initialize VISA resource manager
        try:
            rm = pyvisa.ResourceManager()
        except Exception as e:
            return False

        # Open connection to the N1914A:contentReference[oaicite:17]{index=17}
        self.n1914a = rm.open_resource(visa_address)
        try:
            # Ensure continuous measurement mode on the desired channel
            self.n1914a.write(f":INITiate{channel}:CONTinuous 1")
            # Set the sensor frequency for accurate measurement
            self.n1914a.write(f"SENS:FREQ {int(frequency_hz)}")
            # Query the power measurement (average power in dBm)
        except Exception as e:
            self.n1914a.close()  # Close the instrument connection
            return False
        return True

    def close_n1914a(self):
        self.n1914a.close()  # Close the instrument connection

    def read_n1914a_power(self) -> float:
        power_value = self.n1914a.query_ascii_values(":FETCh:SCALar:POWer:AC?")[0]
        return power_value

    def update_data(self):
        timestamp = time.time()
        #power = self.generate_power_reading()
        power = self.read_n1914a_power()
        self.data.append((timestamp, power))

        # Keep only last 60 readings
        if len(self.data) > 60:
            self.data = self.data[-60:]

        self.update_gui()
        self.root.after(1000, self.update_data)

    def update_gui(self):
        if not self.data:
            return

        # Update current power display
        _, current = self.data[-1]
        self.current_power_var.set(f"{current:.2f} W")

        # Update graph
        self.ax.clear()

        timestamps = [datetime.fromtimestamp(ts) for ts, _ in self.data]
        powers = [p for _, p in self.data]

        # Plot with modern styling
        line = self.ax.plot(timestamps, powers,
                            color='#2c7be5',
                            linewidth=2.5,
                            alpha=0.8,
                            marker='o',
                            markersize=5,
                            markerfacecolor='#ffffff',
                            markeredgecolor='#2c7be5',
                            markeredgewidth=1.5,
                            zorder=3)[0]

        # Fill under line for better visibility
        self.ax.fill_between(timestamps, powers, color='#2c7be5', alpha=0.1)

        # Formatting
        self.ax.set_title('Power Consumption Trend',
                          fontsize=12,
                          fontweight='bold',
                          pad=15,
                          color='#343a40')

        self.ax.set_ylabel('Watts (W)',
                           fontsize=10,
                           labelpad=10,
                           color='#6c757d')

        # Remove x-axis labels completely
        self.ax.set_xticks([])
        self.ax.set_xlabel('')

        # Grid and frame customization
        self.ax.grid(True, linestyle=':', alpha=0.6, color='#e9ecef')
        for spine in ['top', 'right', 'left', 'bottom']:
            self.ax.spines[spine].set_color('#e9ecef')

        # Automatic scaling with padding
        if len(powers) > 0:
            y_padding = max(50, (max(powers) - min(powers)) * 0.2)
            self.ax.set_ylim(
                max(0, min(powers) - y_padding),
                max(powers) + y_padding
            )

        self.figure.tight_layout()
        self.canvas.draw()

    def export_csv(self):
        if not self.data:
            messagebox.showwarning("Warning", "No data to export")
            return

        filename = f"power_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

        try:
            with open(filename, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Timestamp', 'Power (W)'])
                for ts, power in self.data:
                    writer.writerow([
                        datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S'),
                        power
                    ])
            messagebox.showinfo("Success", f"Data exported to {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PowerMonitor(root)
    root.mainloop()