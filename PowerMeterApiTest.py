#!/usr/bin/env python3
"""
PowerMeter API Test Script
Connects to the PowerMeter REST API and displays real-time power readings.
"""

import requests
import time
import json
import sys
from datetime import datetime
import signal

class PowerMeterApiTest:
    def __init__(self, base_url="http://localhost:5000"):
        self.base_url = base_url
        self.running = True
        self.session = requests.Session()
        self.session.timeout = 5
        
        # Setup signal handler for graceful shutdown
        signal.signal(signal.SIGINT, self.signal_handler)
        
    def signal_handler(self, signum, frame):
        """Handle Ctrl+C gracefully"""
        print("\n\nShutting down...")
        self.running = False
        
    def test_connection(self):
        """Test if the API server is reachable"""
        try:
            response = self.session.get(f"{self.base_url}/api/status")
            if response.status_code == 200:
                return True
            else:
                print(f"API server responded with status code: {response.status_code}")
                return False
        except requests.exceptions.ConnectionError:
            print(f"Could not connect to API server at {self.base_url}")
            print("Make sure the PowerMeter application is running with API server enabled.")
            return False
        except Exception as e:
            print(f"Error connecting to API: {e}")
            return False
    
    def get_status(self):
        """Get server status"""
        try:
            response = self.session.get(f"{self.base_url}/api/status")
            if response.status_code == 200:
                return response.json()
            else:
                print(f"Failed to get status: {response.status_code}")
                return None
        except Exception as e:
            print(f"Error getting status: {e}")
            return None
    
    def get_devices(self):
        """Get list of available devices"""
        try:
            response = self.session.get(f"{self.base_url}/api/devices")
            if response.status_code == 200:
                return response.json()
            else:
                print(f"Failed to get devices: {response.status_code}")
                return None
        except Exception as e:
            print(f"Error getting devices: {e}")
            return None
    
    def get_current_power(self):
        """Get current power readings"""
        try:
            response = self.session.get(f"{self.base_url}/api/current")
            if response.status_code == 200:
                return response.json()
            else:
                print(f"Failed to get current power: {response.status_code}")
                return None
        except Exception as e:
            print(f"Error getting current power: {e}")
            return None
    
    def display_status(self, status):
        """Display server status information"""
        print("\n" + "="*60)
        print("POWERMETER API SERVER STATUS")
        print("="*60)
        print(f"Device Connected: {'✓' if status['device_connected'] else '✗'}")
        print(f"Simulation Mode:  {'✓' if status['simulation_mode'] else '✗'}")
        print(f"Monitoring:       {'✓' if status['monitoring'] else '✗'}")
        print(f"Acquisition Freq: {status['acquisition_frequency_ms']} ms")
        print(f"Data Points:      {status['data_points']}")
        print("="*60)
    
    def display_devices(self, devices_response):
        """Display available devices"""
        if not devices_response or not devices_response.get('success'):
            print("Failed to get devices list")
            return
        
        devices = devices_response.get('devices', [])
        if not devices:
            print("No devices found")
            return
        
        print("\n" + "="*60)
        print("AVAILABLE DEVICES")
        print("="*60)
        for i, device in enumerate(devices, 1):
            status = "✓ N1914A" if device['is_n1914a'] else "✗ Other"
            print(f"{i}. {device['resource']}")
            print(f"   Identity: {device['identity']}")
            print(f"   Status:   {status}")
            print()
        print("="*60)
    
    def display_power_reading(self, power_data):
        """Display current power readings"""
        if not power_data:
            print("No power data available")
            return
        
        timestamp = power_data.get('timestamp', 0)
        forward_power = power_data.get('forward_power', 0)
        reflected_power = power_data.get('reflected_power', 0)
        
        # Calculate VSWR if both powers are available
        vswr = "N/A"
        if forward_power > 0 and reflected_power >= 0:
            try:
                reflection_coeff = (reflected_power / forward_power) ** 0.5
                vswr = (1 + reflection_coeff) / (1 - reflection_coeff) if reflection_coeff < 1 else "∞"
                if isinstance(vswr, float):
                    vswr = f"{vswr:.2f}"
            except:
                vswr = "N/A"
        
        # Format timestamp
        try:
            dt = datetime.fromtimestamp(timestamp)
            time_str = dt.strftime("%H:%M:%S")
        except:
            time_str = "N/A"
        
        # Clear line and print power reading
        print(f"\r[{time_str}] Forward: {forward_power:8.2f} W | Reflected: {reflected_power:8.2f} W | VSWR: {vswr:>6}", end="", flush=True)
    
    def run(self):
        """Main test loop"""
        print("PowerMeter API Test Client")
        print("="*60)
        print(f"Connecting to: {self.base_url}")
        print("Press Ctrl+C to quit")
        print("="*60)
        
        # Test connection
        if not self.test_connection():
            return
        
        print("✓ Connected to PowerMeter API server")
        
        # Get and display status
        status = self.get_status()
        if status:
            self.display_status(status)
        
        # Get and display devices
        devices = self.get_devices()
        if devices:
            self.display_devices(devices)
        
        print("\nStarting real-time power monitoring...")
        print("Format: [Time] Forward: XXX.XX W | Reflected: XXX.XX W | VSWR: X.XX")
        print("-" * 80)
        
        # Start monitoring loop
        while self.running:
            try:
                power_data = self.get_current_power()
                self.display_power_reading(power_data)
                time.sleep(1)
                        
            except KeyboardInterrupt:
                break
            except Exception as e:
                print(f"\nError in monitoring loop: {e}")
                time.sleep(1)
        
        print("\n\nTest completed.")

def main():
    import argparse
    
    parser = argparse.ArgumentParser(description='PowerMeter API Test Client')
    parser.add_argument('--url', default='http://localhost:5000', 
                       help='API server URL (default: http://localhost:5000)')
    
    args = parser.parse_args()
    
    test_client = PowerMeterApiTest(args.url)
    test_client.run()

if __name__ == "__main__":
    main() 