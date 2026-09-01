# Spreadsheet-driven PowerPoint builder

Local Python replacement for the Excel VBA / ActiveX presentation builder.
It reads a structured Excel workbook and creates a PowerPoint presentation
directly with `python-pptx`. PowerPoint does not need to be installed to build
the file.

## Main features

- No VBA, ActiveX, COM, or PowerPoint automation
- Spreadsheet remains the source of truth
- Preserves slide hierarchy, object allocation, text runs, paragraph styles,
  shapes, lines, and chart images
- Static browser preview with model/reference comparison
- Simple Windows manager built with Tkinter

## Quick start on Windows

1. Place the required runtime files in the folder structure below.
2. Double-click `START_HERE.bat`.
3. The first run creates a private `.venv` and installs the Python packages.
4. Edit the spreadsheet, refresh the preview, and build the presentation.

Command-line shortcuts:

- `BUILD_NOW.bat` builds the PPTX directly.
- `OPEN_PREVIEW.bat` regenerates and opens the browser preview.

## Required runtime structure

```text
presentation-builder/
├── app.py
├── build_presentation.py
├── preview.py
├── requirements.txt
├── config/
│   └── presentation_config.xlsx
├── assets/
│   └── chart/image files referenced by the workbook
├── reference/
│   └── optional source-slide-01.png, source-slide-02.png, ...
├── preview/
└── output/
```

The workbook must contain these Excel tables:

| Worksheet | Excel table | Purpose |
| --- | --- | --- |
| Control | `tblSettings` | Canvas, output name, and formatting controls |
| Slides | `tblSlides` | Slide-level structure and background |
| Objects | `tblObjects` | Shapes, text boxes, lines, images, and placement |
| Text Runs | `tblTextRuns` | Exact font, size, color, and emphasis |
| Paragraphs | `tblParagraphs` | Alignment and line spacing |
| Formats | `tblFormats` | Simplified reusable styles |

## Direct Python usage

```bash
python -m pip install -r requirements.txt
python build_presentation.py
python preview.py --open
```

The output filename is controlled by `OUTPUT_FILE_NAME` in `tblSettings`.

## Current validation

The USD cash-allocation model used to validate this builder contains seven
slides and 193 objects. The generated deck reproduced all 193 objects and
passed the slide overflow test.

