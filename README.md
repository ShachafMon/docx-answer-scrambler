# docx-answer-scrambler

Shuffles the answer order of multiple-choice questions in a `.docx` exam file
and exports a shuffled PDF with a generated answer key page at the end.

## Download

Grab the latest build for your OS from the [Releases page](https://github.com/ShachafMon/docx-answer-scrambler/releases/latest):

| OS | File |
|---|---|
| Windows | [docx-answer-scrambler.exe](https://github.com/ShachafMon/docx-answer-scrambler/releases/latest/download/docx-answer-scrambler.exe) |
| macOS | [docx-answer-scrambler-mac](https://github.com/ShachafMon/docx-answer-scrambler/releases/latest/download/docx-answer-scrambler-mac) |
| Linux | [docx-answer-scrambler-linux](https://github.com/ShachafMon/docx-answer-scrambler/releases/latest/download/docx-answer-scrambler-linux) |

No Python or installation needed — just download and run.

## Requirements

The app calls two external programs to convert images and export the final PDF.
Install both, then make sure they're on your PATH:

**Windows**
```
winget install Inkscape.Inkscape
winget install TheDocumentFoundation.LibreOffice
```

**macOS**
```
brew install --cask inkscape
brew install --cask libreoffice
```

**Linux (Ubuntu/Debian)**
```
sudo apt install inkscape libreoffice
```

If either is missing, the app will tell you exactly which one and how to install it.

## Usage

1. Put the downloaded executable in the same folder as the `.docx` file(s) you want to shuffle.
2. Run it (double-click, or from a terminal).
3. Find the shuffled PDF(s) in the `shuffled_output` folder created next to it.

Each output PDF has an answer key on its final page.

macOS note: the app isn't signed/notarized, so Gatekeeper may block the first run.
If so: System Settings → Privacy & Security → "Open Anyway", or run
`xattr -d com.apple.quarantine docx-answer-scrambler-mac` in a terminal.

## Building from source

Requires Python 3. Run the build script for your OS — each one creates a
throwaway virtual environment, builds a single-file executable with
PyInstaller, and cleans up after itself:

```
./build_linux.sh      # Linux
./build_mac.sh        # macOS
build_windows.bat     # Windows
```

The executable is written to `dist/`.

Releases are also built automatically by GitHub Actions for all three
platforms whenever a version tag (`v*`) is pushed.
