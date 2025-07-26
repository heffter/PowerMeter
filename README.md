# Power Monitor - Keysight N1914A Integration

A desktop application for displaying and visualizing real-time power measurements from a Keysight N1914A USB power meter. The application provides real-time monitoring, data visualization, and CSV export capabilities with support for both hardware and simulation modes.

## Features

- **Real-time Power Measurement**: Connect to Keysight N1914A via USB and display live power readings
- **Interactive Graphing**: Time-series visualization with a rolling 60-second window
- **Configurable Data Acquisition**: Adjustable acquisition frequency from 100ms to 10 seconds
- **Persistent Configuration**: Save and load device settings between sessions
- **CSV Data Export**: Export measurement data with timestamps and epoch time
- **Simulation Mode**: Test the interface without physical hardware
- **Device Configuration Dialog**: Comprehensive SCPI settings configuration
- **Automatic Device Detection**: Scan and identify compatible USB devices

## Requirements

### Hardware
- Keysight N1914A USB Power Meter (optional - simulation mode available)

### Software
- Python 3.7 or higher
- Keysight IO Libraries Suite (for hardware connection)
- Required Python packages (see requirements.txt)

## Installation

1. **Clone or download this repository**
   ```bash
   git clone <repository-url>
   cd PowerMeter
   ```

2. **Install required Python packages**
   ```bash
   pip install -r requirements.txt
   ```

3. **Install Keysight IO Libraries Suite** (for hardware connection)
   - Download from Keysight website
   - Follow installation instructions for your operating system
   - Verify installation using Keysight Connection Expert

4. **Connect your Keysight N1914A power meter** (optional)
   - Connect via USB cable
   - Verify the device appears in Keysight Connection Expert
   - Note the VISA address for manual configuration if needed

## Usage

### Starting the Application

Run the application from the project directory:

```bash
python PowerMeter.py
```

The application will automatically:
- Load saved configuration (if available)
- Attempt to connect to a Keysight N1914A device
- Fall back to simulation mode if no device is detected
- Display the current power reading and real-time graph

### Main Interface

- **Current Power Display**: Large numerical display showing the current power reading
- **Real-time Graph**: Rolling 60-second window showing power trends
- **Acquisition Frequency Control**: Adjust data collection rate (100ms to 10 seconds)
- **Status Bar**: Shows connection status (Connected/Simulation mode)

### Controls

- **Export Data**: Save the current session's power readings to a CSV file
- **Configure Device**: Open the device configuration dialog
- **Connect/Disconnect Device**: Toggle between real device and simulation mode
- **Exit**: Close the application cleanly

### Device Configuration

The configuration dialog provides comprehensive device setup:

#### Device Connection
- **USB Device Selection**: Scan and select from available devices
- **Manual Connection String**: Enter custom VISA address
- **Test Connection**: Verify device connectivity

#### Measurement Settings
- **Frequency (Hz)**: Set measurement frequency (1 Hz to 50 GHz)
- **Averaging Count**: Configure measurement averaging (1 to 1000)
- **Measurement Unit**: Choose between W, dBm, or dBW
- **Trigger Source**: Select IMM, EXT, or BUS trigger
- **Auto Range**: Enable/disable automatic range selection
- **Manual Range**: Set specific power range (0.1 to 100 W)
- **Integration Time**: Configure measurement integration (0.001 to 1.0 seconds)

### Data Export

Exported CSV files include:
- **Timestamp**: Human-readable date and time
- **Epoch Time**: Unix timestamp for precise time analysis
- **Power (W)**: Measured power value
- **Mode**: Indicates whether data was from real device or simulation

## Configuration

The application automatically creates and maintains a `config.json` file in the same directory as `PowerMeter.py`. This file stores:

- **Device Connection String**: Last successful device connection
- **Measurement Settings**: Frequency, averaging, units, trigger mode, etc.
- **Display Settings**: Update frequency and time window

The configuration is automatically loaded on startup and saved when settings are changed.

## Troubleshooting

### Connection Issues

**Problem**: Device not detected
- **Solution**: Check USB connection and verify device appears in Keysight Connection Expert
- **Alternative**: Use simulation mode for testing

**Problem**: VISA connection errors
- **Solution**: Ensure Keysight IO Libraries Suite is properly installed
- **Solution**: Restart the application and reconnect the device

**Problem**: Invalid SCPI commands
- **Solution**: Verify device model is Keysight N1914A
- **Solution**: Check device configuration in the Configure Device dialog

### Performance Issues

**Problem**: Slow or unresponsive interface
- **Solution**: Reduce acquisition frequency (increase ms value)
- **Solution**: Close other applications to free system resources

**Problem**: Graph not updating
- **Solution**: Check if monitoring is active (should start automatically)
- **Solution**: Verify device connection status

### Data Export Issues

**Problem**: CSV file not created
- **Solution**: Ensure you have write permissions in the selected directory
- **Solution**: Check that data has been collected (graph should show data points)

## Development

### Simulation Mode

The application includes a comprehensive simulation mode for development and testing:

- **Automatic Activation**: Enabled when no hardware is detected
- **Realistic Data**: Generates power readings with realistic fluctuations
- **Full Functionality**: All features work identically to hardware mode
- **Toggle Support**: Can switch between simulation and hardware modes

### Project Structure

```
PowerMeter/
├── PowerMeter.py          # Main application
├── requirements.txt       # Python dependencies
├── README.md             # This documentation
├── config.json           # User configuration (auto-generated)
└── .gitignore           # Git ignore rules
```

### Key Components

- **PowerMonitor Class**: Main application class handling GUI and data acquisition
- **Configuration Management**: JSON-based persistent settings
- **Device Communication**: PyVISA-based hardware interface
- **Data Visualization**: Matplotlib-based real-time graphing
- **Export Functionality**: CSV data export with timestamps

## License

This project is provided as-is for educational and development purposes.

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Verify your hardware and software setup
3. Test with simulation mode to isolate hardware issues
4. Review the configuration settings in the device dialog

