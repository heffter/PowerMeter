#!/bin/bash

echo "Building PowerMeterClientTest..."
dotnet build PowerMeterClientTest.csproj

if [ $? -eq 0 ]; then
    echo "Build successful. Running test..."
    echo "=================================="
    dotnet run --project PowerMeterClientTest.csproj "$@"
else
    echo "Build failed!"
    exit 1
fi 