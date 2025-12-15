# FlightSimTool

A modern .NET 10 WPF application for managing window layouts in flight simulation setups.

## Features
- **Modern UI**: System-aware Dark Mode, rounded controls, and clean aesthetics.
- **Window Capture**: Easily select open windows to create new layouts.
- **Layout Editor**: Fine-tune window positions, sizes, and settings (Border, Fullscreen, etc.) in a dedicated editor tab.
- **Advanced Window Management**:
    - **Follow**: Keep windows positioned relative to their original capture.
    - **Strip Title**: Remove window title bars for clean integration.
    - **Full Screen**: Force windows into "Real Fullscreen" mode (maximized with no borders/caption).
    - **Black Borders**: Add customizable black borders (Letterboxing) to specific windows for bezel correction or projector overlapping.

## Getting Started
1. **Build**: Open the solution in Visual Studio or run `dotnet build` in `src/FlightSimTool`.
2. **Run**: Launch `FlightSimTool.exe`.
3. **Capture**: Go to "Live Windows", check the windows you want to manage, and click "Create New Layout...".
4. **Edit**: Use the "Layout Editor" tab to adjust coordinates, toggle Fullscreen, or add borders.
5. **Apply**: Click "Restore Layout" to move windows, or "Start Overlays" to show black borders.

## Requirements
- Windows 10/11
- .NET 10 Runtime

## Development
- Built with WPF and .NET 10.
- Uses `System.Text.Json` for layout persistence.
- Uses `P/Invoke` (User32, Dwmapi) for advanced window manipulation.


