using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace PowerMeterClientTest
{
    // Data models for API responses
    public class PowerReading
    {
        [JsonProperty("timestamp")]
        public double Timestamp { get; set; }
        
        [JsonProperty("forward_power")]
        public double ForwardPower { get; set; }
        
        [JsonProperty("reflected_power")]
        public double ReflectedPower { get; set; }
        
        [JsonProperty("vswr")]
        public double Vswr { get; set; }
    }

    public class ServerStatus
    {
        [JsonProperty("device_connected")]
        public bool DeviceConnected { get; set; }
        
        [JsonProperty("simulation_mode")]
        public bool SimulationMode { get; set; }
        
        [JsonProperty("monitoring")]
        public bool Monitoring { get; set; }
        
        [JsonProperty("acquisition_frequency_ms")]
        public int AcquisitionFrequencyMs { get; set; }
        
        [JsonProperty("data_points")]
        public int DataPoints { get; set; }
    }

    public class DeviceInfo
    {
        [JsonProperty("resource")]
        public string Resource { get; set; }
        
        [JsonProperty("identity")]
        public string Identity { get; set; }
        
        [JsonProperty("is_n1914a")]
        public bool IsN1914A { get; set; }
    }

    public class DeviceListResponse
    {
        [JsonProperty("success")]
        public bool Success { get; set; }
        
        [JsonProperty("devices")]
        public List<DeviceInfo> Devices { get; set; }
    }

    public class ApiResponse
    {
        [JsonProperty("success")]
        public bool Success { get; set; }
        
        [JsonProperty("message")]
        public string Message { get; set; }
    }

    public class PowerMeterClient : IDisposable
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
                var url = $"{_baseUrl}/api/current";
                Console.WriteLine($"[ASYNC] GET {url}");
                
                var response = await _httpClient.GetAsync(url);
                Console.WriteLine($"[ASYNC] Response Status: {response.StatusCode}");
                
                response.EnsureSuccessStatusCode();

                var json = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"[ASYNC] Response JSON: {json}");
                
                return JsonConvert.DeserializeObject<PowerReading>(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ASYNC] Error getting current power: {ex.Message}");
                return null;
            }
        }

        public PowerReading GetCurrentPower()
        {
            try
            {
                var url = $"{_baseUrl}/api/current";
                Console.WriteLine($"[SYNC] GET {url}");
                
                var response = _httpClient.GetAsync(url).GetAwaiter().GetResult();
                Console.WriteLine($"[SYNC] Response Status: {response.StatusCode}");
                
                response.EnsureSuccessStatusCode();

                var json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                Console.WriteLine($"[SYNC] Response JSON: {json}");
                
                return JsonConvert.DeserializeObject<PowerReading>(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SYNC] Error getting current power: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Get server status
        /// </summary>
        public async Task<ServerStatus> GetStatusAsync()
        {
            try
            {
                var url = $"{_baseUrl}/api/status";
                Console.WriteLine($"[ASYNC] GET {url}");
                
                var response = await _httpClient.GetAsync(url);
                Console.WriteLine($"[ASYNC] Response Status: {response.StatusCode}");
                
                response.EnsureSuccessStatusCode();

                var json = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"[ASYNC] Response JSON: {json}");
                
                return JsonConvert.DeserializeObject<ServerStatus>(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ASYNC] Error getting status: {ex.Message}");
                return null;
            }
        }

        public ServerStatus GetStatus()
        {
            try
            {
                var url = $"{_baseUrl}/api/status";
                Console.WriteLine($"[SYNC] GET {url}");
                
                var response = _httpClient.GetAsync(url).GetAwaiter().GetResult();
                Console.WriteLine($"[SYNC] Response Status: {response.StatusCode}");
                
                response.EnsureSuccessStatusCode();

                var json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                Console.WriteLine($"[SYNC] Response JSON: {json}");
                
                return JsonConvert.DeserializeObject<ServerStatus>(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SYNC] Error getting status: {ex.Message}");
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
                var url = $"{_baseUrl}/api/devices";
                Console.WriteLine($"[ASYNC] GET {url}");
                
                var response = await _httpClient.GetAsync(url);
                Console.WriteLine($"[ASYNC] Response Status: {response.StatusCode}");
                
                response.EnsureSuccessStatusCode();

                var json = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"[ASYNC] Response JSON: {json}");
                
                return JsonConvert.DeserializeObject<DeviceListResponse>(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ASYNC] Error listing devices: {ex.Message}");
                return null;
            }
        }

        public DeviceListResponse ListDevices()
        {
            try
            {
                var url = $"{_baseUrl}/api/devices";
                Console.WriteLine($"[SYNC] GET {url}");
                
                var response = _httpClient.GetAsync(url).GetAwaiter().GetResult();
                Console.WriteLine($"[SYNC] Response Status: {response.StatusCode}");
                
                response.EnsureSuccessStatusCode();

                var json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                Console.WriteLine($"[SYNC] Response JSON: {json}");
                
                return JsonConvert.DeserializeObject<DeviceListResponse>(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SYNC] Error listing devices: {ex.Message}");
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
                var url = $"{_baseUrl}/api/connect";
                Console.WriteLine($"[ASYNC] POST {url}");
                
                var data = new { connection_string = connectionString };
                var json = JsonConvert.SerializeObject(data);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                
                Console.WriteLine($"[ASYNC] Request JSON: {json}");

                var response = await _httpClient.PostAsync(url, content);
                Console.WriteLine($"[ASYNC] Response Status: {response.StatusCode}");
                
                var responseJson = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"[ASYNC] Response JSON: {responseJson}");

                return JsonConvert.DeserializeObject<ApiResponse>(responseJson);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ASYNC] Error connecting to device: {ex.Message}");
                return new ApiResponse { Success = false, Message = ex.Message };
            }
        }

        public ApiResponse ConnectToDevice(string connectionString = "")
        {
            try
            {
                var url = $"{_baseUrl}/api/connect";
                Console.WriteLine($"[SYNC] POST {url}");
                
                var data = new { connection_string = connectionString };
                var json = JsonConvert.SerializeObject(data);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                
                Console.WriteLine($"[SYNC] Request JSON: {json}");

                var response = _httpClient.PostAsync(url, content).GetAwaiter().GetResult();
                Console.WriteLine($"[SYNC] Response Status: {response.StatusCode}");
                
                var responseJson = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                Console.WriteLine($"[SYNC] Response JSON: {responseJson}");

                return JsonConvert.DeserializeObject<ApiResponse>(responseJson);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SYNC] Error connecting to device: {ex.Message}");
                return new ApiResponse { Success = false, Message = ex.Message };
            }
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("PowerMeter C# Client Test Application");
            Console.WriteLine("=====================================\n");

            string baseUrl = "http://localhost:5000";
            if (args.Length > 0)
            {
                baseUrl = args[0];
            }

            Console.WriteLine($"Testing against: {baseUrl}\n");

            using (var client = new PowerMeterClient(baseUrl))
            {
                // Test 1: Server Status (Async)
                Console.WriteLine("=== TEST 1: Server Status (Async) ===");
                var statusAsync = await client.GetStatusAsync();
                if (statusAsync != null)
                {
                    Console.WriteLine($"PASS Async Status: Device={statusAsync.DeviceConnected}, Simulation={statusAsync.SimulationMode}, Monitoring={statusAsync.Monitoring}");
                }
                else
                {
                    Console.WriteLine("FAIL Async Status: Failed to get status");
                }
                Console.WriteLine();

                // Test 2: Server Status (Sync)
                Console.WriteLine("=== TEST 2: Server Status (Sync) ===");
                var statusSync = client.GetStatus();
                if (statusSync != null)
                {
                    Console.WriteLine($"PASS Sync Status: Device={statusSync.DeviceConnected}, Simulation={statusSync.SimulationMode}, Monitoring={statusSync.Monitoring}");
                }
                else
                {
                    Console.WriteLine("FAIL Sync Status: Failed to get status");
                }
                Console.WriteLine();

                // Test 3: Current Power (Async)
                Console.WriteLine("=== TEST 3: Current Power (Async) ===");
                var powerAsync = await client.GetCurrentPowerAsync();
                if (powerAsync != null)
                {
                    Console.WriteLine($"PASS Async Power: Forward={powerAsync.ForwardPower:F2}W, Reflected={powerAsync.ReflectedPower:F2}W, VSWR={powerAsync.Vswr:F2}");
                }
                else
                {
                    Console.WriteLine("FAIL Async Power: Failed to get power reading");
                }
                Console.WriteLine();

                // Test 4: Current Power (Sync)
                Console.WriteLine("=== TEST 4: Current Power (Sync) ===");
                var powerSync = client.GetCurrentPower();
                if (powerSync != null)
                {
                    Console.WriteLine($"PASS Sync Power: Forward={powerSync.ForwardPower:F2}W, Reflected={powerSync.ReflectedPower:F2}W, VSWR={powerSync.Vswr:F2}");
                }
                else
                {
                    Console.WriteLine("FAIL Sync Power: Failed to get power reading");
                }
                Console.WriteLine();

                // Test 5: List Devices (Async)
                Console.WriteLine("=== TEST 5: List Devices (Async) ===");
                var devicesAsync = await client.ListDevicesAsync();
                if (devicesAsync?.Success == true)
                {
                    Console.WriteLine($"PASS Async Devices: Found {devicesAsync.Devices?.Count ?? 0} devices");
                    if (devicesAsync.Devices != null)
                    {
                        foreach (var device in devicesAsync.Devices)
                        {
                            Console.WriteLine($"  - {device.Resource} ({device.Identity})");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("FAIL Async Devices: Failed to get devices");
                }
                Console.WriteLine();

                // Test 6: List Devices (Sync)
                Console.WriteLine("=== TEST 6: List Devices (Sync) ===");
                var devicesSync = client.ListDevices();
                if (devicesSync?.Success == true)
                {
                    Console.WriteLine($"PASS Sync Devices: Found {devicesSync.Devices?.Count ?? 0} devices");
                    if (devicesSync.Devices != null)
                    {
                        foreach (var device in devicesSync.Devices)
                        {
                            Console.WriteLine($"  - {device.Resource} ({device.Identity})");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("FAIL Sync Devices: Failed to get devices");
                }
                Console.WriteLine();

                // Test 7: Connect to Device (Async)
                Console.WriteLine("=== TEST 7: Connect to Device (Async) ===");
                var connectAsync = await client.ConnectToDeviceAsync("");
                if (connectAsync?.Success == true)
                {
                    Console.WriteLine($"PASS Async Connect: {connectAsync.Message}");
                }
                else
                {
                    Console.WriteLine($"FAIL Async Connect: {connectAsync?.Message ?? "Unknown error"}");
                }
                Console.WriteLine();

                // Test 8: Connect to Device (Sync)
                Console.WriteLine("=== TEST 8: Connect to Device (Sync) ===");
                var connectSync = client.ConnectToDevice("");
                if (connectSync?.Success == true)
                {
                    Console.WriteLine($"PASS Sync Connect: {connectSync.Message}");
                }
                else
                {
                    Console.WriteLine($"FAIL Sync Connect: {connectSync?.Message ?? "Unknown error"}");
                }
                Console.WriteLine();

                // Test 9: Continuous Monitoring Test
                Console.WriteLine("=== TEST 9: Continuous Monitoring Test ===");
                Console.WriteLine("Starting continuous monitoring for 10 seconds...");
                Console.WriteLine("Press any key to stop early...");
                
                var startTime = DateTime.Now;
                var monitoringTask = Task.Run(async () =>
                {
                    int count = 0;
                    while (DateTime.Now - startTime < TimeSpan.FromSeconds(10))
                    {
                        try
                        {
                            var reading = await client.GetCurrentPowerAsync();
                            if (reading != null)
                            {
                                count++;
                                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] #{count}: Forward={reading.ForwardPower:F2}W, Reflected={reading.ReflectedPower:F2}W, VSWR={reading.Vswr:F2}");
                            }
                            else
                            {
                                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Failed to get reading");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Error: {ex.Message}");
                        }
                        
                        await Task.Delay(1000);
                        
                        if (Console.KeyAvailable)
                        {
                            Console.ReadKey(true);
                            break;
                        }
                    }
                });

                await monitoringTask;
                Console.WriteLine("Monitoring test completed.\n");
            }

            Console.WriteLine("Test completed. Press any key to exit...");
            Console.ReadKey();
        }
    }
} 