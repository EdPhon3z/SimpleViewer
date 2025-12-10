# Simple Viewer

Current Version: **1.0.0**

![Simple Viewer icon](SimpleViewer/assets/icon.jpg)

Simple Viewer is a minimalist Windows desktop app for quickly browsing photos and videos in a distraction-free, full-screen-friendly viewer.

## Features
- Select a folder or drag one from File Explorer to start browsing instantly
- Filter media by Images, Videos, or both
- Full-screen toggle (F11) hides chrome for immersive viewing
- Slideshow mode with adjustable interval (default 4s)
- Arrow keys navigate forward/backward; Delete moves items into a reserved `_delete_` folder for later review
- Videos loop seamlessly; slideshow timer restarts when navigating manually
- Custom non-commercial license that allows forking and personal use

## Download
Grab the latest `SimpleViewer_win-x64.zip` from the [Releases](https://github.com/EdPhon3z/SimpleViewer/releases). Extract the contents anywhere (e.g., `C:\Apps\SimpleViewer`) and run `SimpleViewer.exe`. No separate .NET install is required.

## Getting Started
1. Ensure the [.NET SDK 8.0](https://dotnet.microsoft.com/en-us/download) or newer is installed.
2. Clone this repository and build the WPF project:
   ```powershell
   git clone <repo-url>
   cd "Simple Viewer/SimpleViewer"
   dotnet build
   ```
3. Run the app:
   ```powershell
   dotnet run
   ```
   or launch the generated `SimpleViewer.exe` under `SimpleViewer/bin/Debug/net8.0-windows/`.

## Usage
- **Select Folder**: Browse to a folder or drag-and-drop it onto the window.
- **Filters**: Toggle Images/Videos checkboxes to limit media types.
- **Slideshow**: Enable slideshow and set interval to auto-advance.
- **Navigation**: Use arrow keys, Delete to move current media into `_delete_`, and F5 to rescan the folder.
- **Full Screen**: Press F11 to toggle chrome and maximize viewing space.

## License
This project is distributed under the [EdPhonez Non-Commercial License (NC-1.0)](https://github.com/EdPhon3z/SimpleViewer/blob/main/LICENSE). Forking and personal/internal use are permitted, but commercial use is prohibited.
