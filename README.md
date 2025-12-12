# Simple Viewer

Current Version: **1.1.0**

![Simple Viewer icon](SimpleViewer/assets/icon.jpg)

Simple Viewer is a minimalist Windows desktop app for quickly browsing photos and videos in a distraction-free, full-screen-friendly viewer.

## Features
- Select a folder or drag one from File Explorer to start browsing instantly
- Filter media by Images, Videos, or both
- Full-screen toggle (F11) hides chrome for immersive viewing
- Slideshow mode with adjustable interval (default 4s)
- Remembers the last five folders you visited and exposes them via the new Recent button
- Arrow keys navigate forward/backward; Delete moves items into a reserved `_delete_` folder for later review
- Videos loop seamlessly; slideshow timer restarts when navigating manually
- Custom non-commercial license that allows forking and personal use

## Screenshots
| Single View (toolbar + metadata) | Recent Folders flyout |
| --- | --- |
| ![Single view screenshot](SimpleViewer/assets/ScreenshotOptions.png) | ![Recent folders screenshot](SimpleViewer/assets/ScreenshotRecent.png) |

| Grid View (standard) | Grid View (zoomed thumbnails) |
| --- | --- |
| ![Grid view screenshot](SimpleViewer/assets/ScreenshotGrid.png) | ![Grid view zoomed screenshot](SimpleViewer/assets/ScreenshotSmallGrid.png) |

| Video Playback with Pause Overlay |
| --- |
| ![Video pause screenshot](SimpleViewer/assets/ScreenshotVideoPauseVideo.png) |

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

## Usage & Controls
- **Select Folder**: Browse to a folder or drag-and-drop it onto the window.
- **Recent Folders**: Use the clock/bookmark button to reopen the last five folders or clear the list.
- **Filters**: Toggle Images/Videos checkboxes to limit media types.
- **Slideshow**: Enable slideshow and set interval to auto-advance.
- **Navigation**: Arrow Left/Right (and Up/Down for images) or the mouse wheel move through items; Delete moves current media into `_delete_`; F5 rescans the folder.
- **Full Screen**: Press F11 to toggle chrome and maximize viewing space; ESC exits full screen (or closes the app if already windowed).
- **Zooming/Panning**: Ctrl + mouse wheel zooms images in/out; press `0` (or NumPad `0`) to reset to 100%; click-drag pans when zoomed past 100%.
- **Grid Zoom**: While in Grid view, Ctrl + mouse wheel scales thumbnail size for quick browsing.
- **Video Audio**: While a video is playing, Up increases and Down decreases volume; press `Space` to pause/resume and `M` to mute/unmute.
- **Grid View**: Press `G` or use Options → View → Grid to browse thumbnail tiles; double-click a tile to jump back into single view.
- **Sorting**: Options → Sort by lets you switch between newest, modified, alphabetical, or random ordering of items.
- **Metadata Panel**: Press `X` to toggle EXIF/Comfy metadata with a copy button for the full text.

## License
This project is distributed under the [EdPhonez Non-Commercial License (NC-1.0)](https://github.com/EdPhon3z/SimpleViewer/blob/main/LICENSE). Forking and personal/internal use are permitted, but commercial use is prohibited.
