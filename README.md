# MediaSort

A sleek, dark-themed desktop application for automatically sorting and organizing media files into categorized sub-folders. Built with Python and tkinter.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

## Features

- **Four Media Types**: Sort images, videos, documents, and audio files
- **Smart Classification**: Automatically categorizes files based on content, dimensions, and metadata
- **Dark Modern UI**: Clean, professional interface with color-coded categories
- **Non-Destructive**: Choose to copy or move files into organized sub-folders
- **Recursive Scanning**: Optionally scan sub-folders for scattered files
- **Auto-Organize**: Automatically move files after scanning (optional)
- **Real-time Progress**: Live progress bar and status updates during operations
- **Safe Naming**: Handles duplicate filenames and forbidden characters automatically

## Installation

### Requirements

- Python 3.8 or higher
- Pillow (`pip install Pillow`)
- ffprobe (for video dimension detection) OR opencv-python (`pip install opencv-python`)

### Quick Start

```bash
# Clone or download MediaSort.py
# Install dependencies
pip install Pillow opencv-python

# Run the application
python MediaSort.py
```

> **Note**: `ffprobe` is preferred for video analysis. If unavailable, the app falls back to OpenCV automatically.

## Usage

1. **Launch** the application
2. **Select** a folder to scan using the Browse button
3. **Choose** your media type tab (Images, Videos, Documents, or Audio)
4. **Configure** options:
   - ☑ Scan sub-folders recursively
   - ☑ Auto-move files after scan
5. **Click** ▶ Scan to analyze files
6. **Review** results in the categorized tabs and summary cards
7. **Click** Move or Copy to organize files into sub-folders

## Classification Rules

### Images
| Category | Criteria | Color |
|----------|----------|-------|
| Wallpapers | Filename contains "wallhaven" | Purple |
| Screenshots | Filename contains "screenshot" | Teal |
| WebP Images | `.webp` extension | Amber |
| Square Images | Width/height ratio within 5% of 1:1 | Cyan |
| Other Images | Everything else | Rose |

### Videos
| Category | Criteria | Color |
|----------|----------|-------|
| Screen Recordings | Filename starts with "screen_recording" | Green |
| 16:9 Landscape | Aspect ratio ≈ 16:9 (±5%) | Sky Blue |
| 9:16 Portrait | Aspect ratio ≈ 9:16 (±5%) | Pink |
| Other | Everything else | Peach |

### Documents
| Category | Extensions | Color |
|----------|------------|-------|
| Word Documents | `.doc` `.docx` `.odt` `.rtf` ... | Light Blue |
| Spreadsheets | `.xls` `.xlsx` `.ods` `.csv` `.tsv` ... | Light Green |
| Presentations | `.ppt` `.pptx` `.odp` ... | Yellow |
| PDFs | `.pdf` | Rose |
| Text & Markdown | `.txt` `.md` `.rst` `.log` ... | Lilac |
| Code & Scripts | `.py` `.js` `.html` `.css` `.json` `.xml` ... | Cyan |
| Archives | `.zip` `.rar` `.7z` `.tar` `.gz` ... | Orange |
| Other Documents | Everything else | Blue-Grey |

### Audio
| Category | Extensions | Color |
|----------|------------|-------|
| Lossless | `.flac` `.wav` `.aiff` `.alac` `.ape` ... | Lime |
| Compressed | `.mp3` `.aac` `.ogg` `.opus` `.m4a` ... | Teal |
| Playlists | `.m3u` `.m3u8` `.pls` `.xspf` ... | Purple |
| Other Audio | Everything else | Pink |

## Screenshots

*The application features a dark-themed interface with:*
- Color-coded category cards showing file counts
- Tabbed file listings with detailed metadata
- Progress tracking and status bar
- Clean, modern button controls

## File Structure

```
YourFolder/
├── MediaSort.py          # Main application
└── YourScannedFolder/
    ├── Wallpapers/         # Sorted sub-folders created here
    ├── Screenshots/
    ├── 16:9 Landscape/
    ├── Word Documents/
    ├── PDFs/
    ├── Lossless/
    └── ...
```

## Technical Details

- **GUI Framework**: tkinter (stdlib) with custom dark theme
- **Image Processing**: Pillow for dimension extraction
- **Video Processing**: ffprobe (preferred) or OpenCV fallback
- **Threading**: Background scanning and file operations to keep UI responsive
- **Cross-Platform**: Works on Windows, macOS, and Linux

## Safety Features

- Duplicate filename detection with automatic renaming (`file_1.ext`, `file_2.ext`)
- Invalid character stripping for folder names
- Source and destination collision detection
- Non-destructive copy option available
- Cancel scan mid-operation

## License

MIT License — feel free to use, modify, and distribute.

## Contributing

Contributions welcome! The codebase is modular — each media type is self-contained in its own panel class, making it easy to add new categories or file types.
