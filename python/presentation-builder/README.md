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

## Corporate environment

This version does not use batch files, ActiveX, COM, or PowerPoint automation.
It also does not download or install packages automatically.

Use a corporate-approved Python 3.10+ installation. Ask your Python or desktop
administrator to make the packages in `requirements.txt` available through the
company's approved Python environment or internal package index.

From PowerShell, Command Prompt, an IDE terminal, or the approved Python console:

```text
python run.py status
python run.py build
python run.py preview
python run.py preview --open
```

Run `python run.py` without a command to use the interactive text menu. The
optional Tkinter interface remains available through `python run.py gui`.

## Required runtime structure

```text
presentation-builder/
├── app.py
├── build_presentation.py
├── preview.py
├── run.py
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

## Dependency setup

Use the method approved by the company. If the corporate Python environment
permits pip and already points to an internal package index:

```text
python -m pip install -r requirements.txt
```

No administrator access is required by the code itself. A virtual environment
is optional and is not created automatically.

The output filename is controlled by `OUTPUT_FILE_NAME` in `tblSettings`.

## Current validation

The USD cash-allocation model used to validate this builder contains seven
slides and 193 objects. The generated deck reproduced all 193 objects and
passed the slide overflow test.
