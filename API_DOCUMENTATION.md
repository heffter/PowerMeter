# PowerMeter REST API Documentation

The PowerMeter application now includes an integrated REST API server that provides access to power measurement data and device control functionality.

## Configuration

The API server can be configured through the PowerMeter GUI:

1. Click "Configure Device" button
2. Navigate to the "API Server Configuration" section
3. Enable/disable the API server
4. Set the port (1024-65535)
5. Set the host (0.0.0.0 for all interfaces, 127.0.0.1 for localhost only)

## API Endpoints

### GET /api/status
Get the current status of the power meter.

**Response:**
```json
{
    "device_connected": true,
    "simulation_mode": false,
    "monitoring": true,
    "acquisition_frequency_ms": 1000,
    "data_points": 45
}
```

### GET /api/current
Get the current power readings.

**Response:**
```json
{
    "timestamp": 1703123456.789,
    "forward_power": 750.25,
    "reflected_power": 45.12
}
```

### GET /api/history?limit=100
Get power history data.

**Parameters:**
- `limit` (optional): Number of data points to return (default: 100, max: available data points)

**Response:**
```json
[
    {
        "timestamp": 1703123456.789,
        "forward_power": 750.25,
        "reflected_power": 45.12
    },
    {
        "timestamp": 1703123457.789,
        "forward_power": 751.30,
        "reflected_power": 44.98
    }
]
```

### GET /api/devices
List available VISA devices.

**Response:**
```json
{
    "success": true,
    "devices": [
        {
            "resource": "USB0::0x0957::0x0607::MY12345678::INSTR",
            "identity": "Keysight Technologies,N1914A,MY12345678,A.01.01",
            "is_n1914a": true
        }
    ]
}
```

## C# Client Example

```csharp
using System;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;

public class PowerMeterClient
{
    private readonly HttpClient _httpClient;
    private readonly string _baseUrl;

    public PowerMeterClient(string baseUrl = "http://localhost:5000")
    {
        _baseUrl = baseUrl;
        _httpClient = new HttpClient();
    }

    public async Task<PowerReading> GetCurrentPowerAsync()
    {
        var response = await _httpClient.GetAsync($"{_baseUrl}/api/current");
        response.EnsureSuccessStatusCode();
        
        var json = await response.Content.ReadAsStringAsync();
        return JsonConvert.DeserializeObject<PowerReading>(json);
    }

    public async Task<ServerStatus> GetStatusAsync()
    {
        var response = await _httpClient.GetAsync($"{_baseUrl}/api/status");
        response.EnsureSuccessStatusCode();
        
        var json = await response.Content.ReadAsStringAsync();
        return JsonConvert.DeserializeObject<ServerStatus>(json);
    }
}

public class PowerReading
{
    public double Timestamp { get; set; }
    public double ForwardPower { get; set; }
    public double ReflectedPower { get; set; }
}

public class ServerStatus
{
    public bool DeviceConnected { get; set; }
    public bool SimulationMode { get; set; }
    public bool Monitoring { get; set; }
    public int AcquisitionFrequencyMs { get; set; }
    public int DataPoints { get; set; }
}
```

## Usage Examples

### Get Current Power Reading
```bash
curl http://localhost:5000/api/current
```

### Get Server Status
```bash
curl http://localhost:5000/api/status
```

### Get Power History (last 50 points)
```bash
curl http://localhost:5000/api/history?limit=50
```

### List Available Devices
```bash
curl http://localhost:5000/api/devices
```

## Features

- **Real-time Data**: Access current power readings and historical data
- **Device Information**: List available VISA devices
- **Status Monitoring**: Check device connection and monitoring status
- **CORS Enabled**: Cross-origin requests supported for web applications
- **Thread-safe**: API server runs in a separate thread
- **Configuration Persistence**: API settings saved to config.json

## Security Notes

- The API server is designed for local network use
- Use appropriate firewall rules for production deployment
- Consider using HTTPS for sensitive environments
- The API provides read-only access to power data and device information

## Troubleshooting

1. **API Server Won't Start**: Check if the port is already in use
2. **Connection Refused**: Verify the API server is enabled and running
3. **No Data**: Ensure the power meter is connected and monitoring is active
4. **CORS Issues**: The API includes CORS headers for web applications 