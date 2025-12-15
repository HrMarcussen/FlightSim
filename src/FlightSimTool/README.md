# FlightSim Layout Tool (WPF)

This is a modern C# WPF replacement for the PowerShell window layout scripts.

## Prerequisites
- **.NET 6.0 SDK** (or later)
  - Download: [https://dotnet.microsoft.com/download](https://dotnet.microsoft.com/download)

## How to Build
Open a terminal in this directory (`src/FlightSimTool`) and run:

```powershell
dotnet build
```

## How to Run
After building, you can run the app directly:

```powershell
dotnet run
```

Or execute the generated `.exe` in `bin\Debug\net6.0-windows\FlightSimTool.exe`.

## Migration from PowerShell
- This app uses the same `WindowLayout.json` format (partially compatible).
- It provides a visual **Dashboard** to replace `WindowLayout.ps1`.
- **Overlays** are now handled internally by the app (no separate scripts).
