using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace PowerMeterClient
{
    // Data models for API responses
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

    public class DeviceInfo
    {
        public string Resource { get; set; }
        public string Identity { get; set; }
        public bool IsN1914A { get; set; }
    }

    public class DeviceListResponse
    {
        public bool Success { get; set; }
        public List<DeviceInfo> Devices { get; set; }
    }

    public class ApiResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; }
    }

    public class PowerMeterClient
    {
        private readonly HttpClient _httpClient;
        private readonly string _baseUrl;

        public PowerMeterClient(string baseUrl = "http://localhost:5000")
        {
            _baseUrl = baseUrl;
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromSeconds(10);
        }

        /// <summary>
        /// Get current power readings
        /// </summary>
        public async Task<PowerReading> GetCurrentPowerAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync($"{_baseUrl}/api/current");
                response.EnsureSuccessStatusCode();
                
                var json = await response.Content.ReadAsStringAsync();
                return JsonConvert.DeserializeObject<PowerReading>(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting current power: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Get power history data
        /// </summary>
        public async Task<List<PowerReading>> GetPowerHistoryAsync(int limit = 100)
        {
            try
            {
                var response = await _httpClient.GetAsync($"{_baseUrl}/api/history?limit={limit}");
                response.EnsureSuccessStatusCode();
                
                var json = await response.Content.ReadAsStringAsync();
                return JsonConvert.DeserializeObject<List<PowerReading>>(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting power history: {ex.Message}");
                return new List<PowerReading>();
            }
        }

        /// <summary>
        /// Get server status
        /// </summary>
        public async Task<ServerStatus> GetStatusAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync($"{_baseUrl}/api/status");
                response.EnsureSuccessStatusCode();
                
                var json = await response.Content.ReadAsStringAsync();
                return JsonConvert.DeserializeObject<ServerStatus>(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting status: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// List available devices
        /// </summary>
        public async Task<DeviceListResponse> ListDevicesAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync($"{_baseUrl}/api/devices");
                response.EnsureSuccessStatusCode();
                
                var json = await response.Content.ReadAsStringAsync();
                return JsonConvert.DeserializeObject<DeviceListResponse>(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing devices: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Connect to a specific device
        /// </summary>
        public async Task<ApiResponse> ConnectToDeviceAsync(string connectionString = "")
        {
            try
            {
                var data = new { connection_string = connectionString };
                var json = JsonConvert.SerializeObject(data);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                
                var response = await _httpClient.PostAsync($"{_baseUrl}/api/connect", content);
                var responseJson = await response.Content.ReadAsStringAsync();
                
                return JsonConvert.DeserializeObject<ApiResponse>(responseJson);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error connecting to device: {ex.Message}");
                return new ApiResponse { Success = false, Message = ex.Message };
            }
        }

        /// <summary>
        /// Toggle simulation mode
        /// </summary>
        public async Task<ApiResponse> ToggleSimulationAsync(bool enable = true)
        {
            try
            {
                var data = new { enable = enable };
                var json = JsonConvert.SerializeObject(data);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                
                var response = await _httpClient.PostAsync($"{_baseUrl}/api/simulation", content);
                var responseJson = await response.Content.ReadAsStringAsync();
                
                return JsonConvert.DeserializeObject<ApiResponse>(responseJson);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error toggling simulation: {ex.Message}");
                return new ApiResponse { Success = false, Message = ex.Message };
            }
        }

        /// <summary>
        /// Set acquisition frequency
        /// </summary>
        public async Task<ApiResponse> SetAcquisitionFrequencyAsync(int frequencyMs)
        {
            try
            {
                var data = new { frequency_ms = frequencyMs };
                var json = JsonConvert.SerializeObject(data);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                
                var response = await _httpClient.PostAsync($"{_baseUrl}/api/frequency", content);
                var responseJson = await response.Content.ReadAsStringAsync();
                
                return JsonConvert.DeserializeObject<ApiResponse>(responseJson);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error setting frequency: {ex.Message}");
                return new ApiResponse { Success = false, Message = ex.Message };
            }
        }

        /// <summary>
        /// Start continuous monitoring of power readings
        /// </summary>
        public async Task StartMonitoringAsync(Action<PowerReading> onPowerUpdate, int intervalMs = 1000)
        {
            while (true)
            {
                try
                {
                    var powerReading = await GetCurrentPowerAsync();
                    if (powerReading != null)
                    {
                        onPowerUpdate(powerReading);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error in monitoring loop: {ex.Message}");
                }
                
                await Task.Delay(intervalMs);
            }
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }

    // Example usage
    class Program
    {
        static async Task Main(string[] args)
        {
            using var client = new PowerMeterClient("http://localhost:5000");
            
            Console.WriteLine("PowerMeter C# Client Example");
            Console.WriteLine("============================\n");

            // Get server status
            var status = await client.GetStatusAsync();
            if (status != null)
            {
                Console.WriteLine($"Server Status:");
                Console.WriteLine($"  Device Connected: {status.DeviceConnected}");
                Console.WriteLine($"  Simulation Mode: {status.SimulationMode}");
                Console.WriteLine($"  Monitoring: {status.Monitoring}");
                Console.WriteLine($"  Acquisition Frequency: {status.AcquisitionFrequencyMs}ms");
                Console.WriteLine($"  Data Points: {status.DataPoints}\n");
            }

            // Get current power reading
            var currentPower = await client.GetCurrentPowerAsync();
            if (currentPower != null)
            {
                Console.WriteLine($"Current Power Reading:");
                Console.WriteLine($"  Timestamp: {DateTimeOffset.FromUnixTimeSeconds((long)currentPower.Timestamp)}");
                Console.WriteLine($"  Forward Power: {currentPower.ForwardPower:F2} W");
                Console.WriteLine($"  Reflected Power: {currentPower.ReflectedPower:F2} W");
                Console.WriteLine($"  Simulation Mode: {currentPower.SimulationMode}\n");
            }

            // List available devices
            var devices = await client.ListDevicesAsync();
            if (devices?.Success == true)
            {
                Console.WriteLine("Available Devices:");
                foreach (var device in devices.Devices)
                {
                    Console.WriteLine($"  {device.Resource} - {device.Identity} (N1914A: {device.IsN1914A})");
                }
                Console.WriteLine();
            }

            // Example of continuous monitoring
            Console.WriteLine("Starting continuous monitoring (press any key to stop)...");
            var cts = new System.Threading.CancellationTokenSource();
            
            var monitoringTask = Task.Run(async () =>
            {
                await client.StartMonitoringAsync(powerReading =>
                {
                    Console.WriteLine($"[{DateTimeOffset.FromUnixTimeSeconds((long)powerReading.Timestamp):HH:mm:ss}] " +
                                    $"Forward: {powerReading.ForwardPower:F2}W, " +
                                    $"Reflected: {powerReading.ReflectedPower:F2}W");
                }, 2000); // Update every 2 seconds
            }, cts.Token);

            Console.ReadKey();
            cts.Cancel();
            
            Console.WriteLine("\nMonitoring stopped.");
        }
    }
} 