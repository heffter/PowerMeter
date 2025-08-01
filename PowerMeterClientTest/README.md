# PowerMeter Client Test

This is a comprehensive test application for the PowerMeterClient class, designed to test both synchronous and asynchronous operations with detailed logging.

## Requirements

- Windows 8.1 or later
- Visual Studio 2019 (or later)
- .NET Framework 4.7.2 (included with Visual Studio 2019)

## Setup Instructions

1. **Open the Solution**
   - Open `PowerMeterClientTest.sln` in Visual Studio 2019
   - The solution contains the PowerMeterClientTest project

2. **Restore NuGet Packages**
   - Right-click on the solution in Solution Explorer
   - Select "Restore NuGet Packages"
   - This will download Newtonsoft.Json dependency

3. **Build the Project**
   - Press Ctrl+Shift+B or go to Build → Build Solution
   - Ensure the build completes successfully

## Running the Test

### Method 1: Visual Studio
1. Press F5 or go to Debug → Start Debugging
2. The application will run and show detailed test results
3. Press any key to exit when testing is complete

### Method 2: Command Line
1. Open Command Prompt in the project directory
2. Run: `dotnet run`
3. Or run the executable directly: `bin\Debug\net472\PowerMeterClientTest.exe`

## Test Features

The test application includes:

- **Detailed Logging**: Every API call logs URL, status code, and JSON response
- **Sync/Async Testing**: Tests both synchronous and asynchronous versions
- **Comprehensive Coverage**: Tests all major API endpoints
- **Continuous Monitoring**: 10-second real-time monitoring test
- **Error Handling**: Detailed error reporting

## API Endpoints Tested

1. **Server Status** (`/api/status`)
2. **Current Power** (`/api/current`)
3. **Device List** (`/api/devices`)
4. **Device Connection** (`/api/connect`)

## Customization

You can modify the base URL by:
- Changing the default URL in the code
- Passing command line arguments
- Editing the `baseUrl` variable in the `Main` method

## Troubleshooting

### Common Issues

1. **Connection Refused**
   - Ensure the PowerMeter API server is running
   - Check if the server is on the correct port (default: 5000)

2. **Build Errors**
   - Ensure .NET Framework 4.7.2 is installed
   - Restore NuGet packages if Newtonsoft.Json is missing

3. **Runtime Errors**
   - Check firewall settings
   - Verify the API server is accessible from the test machine

### Debug Information

The test application provides extensive logging:
- `[SYNC]` or `[ASYNC]` prefixes indicate the operation type
- Full HTTP request URLs are logged
- Response status codes and JSON content are displayed
- Error messages include detailed exception information

## File Structure

```
PowerMeterClientTest/
├── PowerMeterClientTest.sln          # Visual Studio solution file
├── PowerMeterClientTest.csproj       # Project file
├── PowerMeterClientTest.cs           # Main test application
├── README.md                         # This file
└── run_test.sh                       # Linux build script (for reference)
``` 