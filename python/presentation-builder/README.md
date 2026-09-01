# Presentation Model Manager

Corporate-safe, spreadsheet-driven PowerPoint builder contained in one Python
file: `presentation_manager.py`.

The file creates the standard model workbook structure, manages multiple model
profiles, validates tables and assets, builds PPTX presentations, generates
browser previews, and includes optional text and Tkinter managers.

It does not use BAT files, VBA, ActiveX, COM, PowerPoint automation, or automatic
package downloads.

## Corporate requirements

- Corporate-approved Python 3.10+
- `python-pptx`
- `openpyxl`
- `Pillow`

Install dependencies only through the company's approved Python environment or
internal package index.

## Commands

```text
python presentation_manager.py list
python presentation_manager.py create new_model
python presentation_manager.py clone source_model scenario_model
python presentation_manager.py validate usd_cash_allocation
python presentation_manager.py build usd_cash_allocation
python presentation_manager.py preview usd_cash_allocation --open
python presentation_manager.py gui
```

Running `python presentation_manager.py` without a command opens the interactive
text manager.

## Model structure

```text
models/
└── MODEL_NAME/
    ├── model.xlsx
    ├── assets/
    └── reference/
```

Each `model.xlsx` contains these managed Excel tables:

| Worksheet | Table | Purpose |
| --- | --- | --- |
| Control | `tblSettings` | Canvas, output name, and formatting controls |
| Formats | `tblFormats` | Reusable text hierarchy and styles |
| Slides | `tblSlides` | Slide-level settings |
| Objects | `tblObjects` | Shapes, text, charts, geometry, and allocation |
| Text Runs | `tblTextRuns` | Exact font formatting |
| Paragraphs | `tblParagraphs` | Paragraph alignment and spacing |
| Assets | `tblAssets` | Chart and image inventory |

Generated presentations are written to `output/MODEL_NAME/`. Browser previews
are written to `preview/MODEL_NAME/`.
