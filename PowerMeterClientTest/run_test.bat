@echo off
echo Building PowerMeterClientTest...
dotnet build PowerMeterClientTest.csproj

if %ERRORLEVEL% EQU 0 (
    echo Build successful. Running test...
    echo ==================================
    dotnet run --project PowerMeterClientTest.csproj %*
) else (
    echo Build failed!
    pause
    exit /b 1
)

pause 