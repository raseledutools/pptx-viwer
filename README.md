# PPTX Viewer

A modern WPF-based PowerPoint (.pptx) viewer for Windows — no Office required.

## Features
- 📂 Open PPTX files (drag & drop supported)
- 🖼️ Slide thumbnail sidebar
- ⛶ Fullscreen slideshow mode (F5)
- 📝 Speaker notes panel
- 🔍 Zoom in/out (+ / - keys)
- ⌨️ Keyboard navigation (← →, Page Up/Down, Home/End)
- 🖱️ Click thumbnail to jump to slide

## Keyboard Shortcuts
| Key | Action |
|-----|--------|
| `→` / `Space` / `Page Down` | Next slide |
| `←` / `Page Up` | Previous slide |
| `Home` | First slide |
| `End` | Last slide |
| `F5` | Fullscreen slideshow |
| `ESC` | Exit fullscreen |
| `+` / `-` | Zoom in/out |

## Build (GitHub Actions)

Push to `main` → GitHub Actions builds automatically → download `PptxViewer.exe` from Artifacts.

### Release build
```bash
git tag v1.0.0
git push origin v1.0.0
```
This triggers a GitHub Release with the `.exe` attached.

## Local build
```bash
dotnet publish PptxViewer/PptxViewer.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o ./publish
```

## Requirements
- .NET 8 SDK (build only)
- Windows 7+ (runtime, no install needed — self-contained EXE)
