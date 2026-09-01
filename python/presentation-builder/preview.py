from __future__ import annotations

import argparse
import html
import json
import os
import webbrowser
from collections import defaultdict
from pathlib import Path

from build_presentation import DEFAULT_CONFIG, ROOT, load_model


PREVIEW_DIR = ROOT / "preview"
PREVIEW_FILE = PREVIEW_DIR / "index.html"


def _color(value, transparency=0, fallback="transparent") -> str:
    if not value or str(value).upper() == "NONE" or float(transparency or 0) >= 0.995:
        return fallback
    return str(value)


def _text_html(row, slide_row, runs_by_object, formats_by_id) -> str:
    object_id = str(row.get("Object_ID"))
    exact = runs_by_object.get(object_id, {})
    fallback = formats_by_id.get(str(row.get("Format_ID")), {})
    dark = str(slide_row.get("Theme_Mode") or "").lower() == "dark"
    fallback_color = fallback.get("Font_Color_Dark_Page" if dark else "Font_Color_Light_Page") or "#0B1826"
    output = []
    for paragraph_no, line in enumerate(str(row.get("Text_Content") or "").split("\n"), start=1):
        runs = exact.get(paragraph_no, [])
        if runs:
            spans = []
            for run in runs:
                style = (
                    f"font-family:{html.escape(str(run.get('Font_Family') or 'Aptos'))};"
                    f"font-size:{float(run.get('Font_Size_pt') or 12)}pt;"
                    f"color:{html.escape(str(run.get('Font_Color') or '#0B1826'))};"
                    f"font-weight:{'700' if run.get('Bold') else '400'};"
                    f"font-style:{'italic' if run.get('Italic') else 'normal'};"
                )
                spans.append(f"<span style=\"{style}\">{html.escape(str(run.get('Run_Text') or ''))}</span>")
            output.append(f"<div>{''.join(spans)}</div>")
        elif line:
            style = (
                f"font-family:{html.escape(str(fallback.get('Font_Family') or 'Aptos'))};"
                f"font-size:{float(fallback.get('Font_Size_pt') or 12)}pt;"
                f"color:{html.escape(str(fallback_color))};"
                f"font-weight:{'700' if fallback.get('Bold') else '400'};"
            )
            output.append(f"<div style=\"{style}\">{html.escape(line)}</div>")
        else:
            output.append("<div>&nbsp;</div>")
    return "".join(output)


def generate_preview(config_path: Path = DEFAULT_CONFIG) -> Path:
    model = load_model(config_path)
    slide_rows = sorted(model["slides"], key=lambda row: int(row["Slide_No"]))
    objects_by_slide = defaultdict(list)
    for row in sorted(model["objects"], key=lambda item: (int(item["Slide_No"]), int(item["Z_Order"]))):
        if bool(row.get("Enabled")):
            objects_by_slide[int(row["Slide_No"])].append(row)
    runs_by_object = defaultdict(lambda: defaultdict(list))
    for row in model["runs"]:
        runs_by_object[str(row["Object_ID"])][int(row["Paragraph_No"])].append(row)
    formats_by_id = {str(row["Format_ID"]): row for row in model["formats"]}

    slide_cards = []
    for slide_row in slide_rows:
        slide_no = int(slide_row["Slide_No"])
        elements = []
        for row in objects_by_slide[slide_no]:
            left, top = float(row.get("Left_px") or 0), float(row.get("Top_px") or 0)
            width, height = float(row.get("Width_px") or 0), float(row.get("Height_px") or 0)
            common = f"left:{left}px;top:{top}px;width:{width}px;height:{height}px;"
            object_type = str(row.get("Object_Type") or "")
            geometry = str(row.get("Geometry") or "rect")
            if object_type in {"Image", "ChartImage"}:
                asset = Path(str(row.get("Asset_ID") or "")).name
                elements.append(f"<img class=\"obj\" style=\"{common}\" src=\"../assets/{html.escape(asset)}\" alt=\"chart\">")
            elif object_type == "Line" or geometry == "line":
                stroke = _color(row.get("Line_Color"), row.get("Line_Transparency"), "transparent")
                line_width = max(float(row.get("Line_Width_px") or 1), 0.5)
                elements.append(
                    f"<svg class=\"obj line\" style=\"left:{left}px;top:{top}px;width:{max(width,1)}px;height:{max(abs(height),1)}px\">"
                    f"<line x1=\"0\" y1=\"0\" x2=\"{width}\" y2=\"{height}\" stroke=\"{stroke}\" stroke-width=\"{line_width}\"/></svg>"
                )
            else:
                fill = _color(row.get("Fill_Color"), row.get("Fill_Transparency"), "transparent")
                border = _color(row.get("Line_Color"), row.get("Line_Transparency"), "transparent")
                border_width = float(row.get("Line_Width_px") or 0)
                radius = "10px" if geometry == "roundRect" else "0"
                elements.append(
                    f"<div class=\"obj\" style=\"{common}background:{fill};border:{border_width}px solid {border};border-radius:{radius};box-sizing:border-box\"></div>"
                )
            if str(row.get("Text_Content") or ""):
                vertical = {"middle": "center", "bottom": "flex-end"}.get(str(row.get("Vertical_Alignment") or "top").lower(), "flex-start")
                align = str(row.get("Text_Alignment") or "left").lower()
                margin = (
                    float(row.get("Margin_Top_px") or 0),
                    float(row.get("Margin_Right_px") or 0),
                    float(row.get("Margin_Bottom_px") or 0),
                    float(row.get("Margin_Left_px") or 0),
                )
                text_style = (
                    f"{common}display:flex;align-items:{vertical};text-align:{align};"
                    f"padding:{margin[0]}px {margin[1]}px {margin[2]}px {margin[3]}px;"
                    "box-sizing:border-box;overflow:hidden;line-height:1;"
                )
                elements.append(f"<div class=\"obj text\" style=\"{text_style}\"><div style=\"width:100%\">{_text_html(row, slide_row, runs_by_object, formats_by_id)}</div></div>")

        reference = f"../reference/source-slide-{slide_no:02d}.png"
        slide_cards.append(
            f"<section class=\"card\"><header><strong>Slide {slide_no}</strong><button onclick=\"toggleView(this)\">Show reference</button></header>"
            f"<div class=\"viewport model\"><div class=\"slide\" style=\"background:{slide_row.get('Background_Color') or '#071B27'}\">{''.join(elements)}</div></div>"
            f"<div class=\"viewport reference hidden\"><img src=\"{reference}\" alt=\"reference slide {slide_no}\"></div></section>"
        )

    page = f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Presentation Builder Preview</title>
<style>
body{{margin:0;background:#eef2f5;color:#0b1826;font:14px Arial,sans-serif}}.top{{position:sticky;top:0;z-index:9;background:#071b27;color:white;padding:18px 28px;display:flex;justify-content:space-between;align-items:center}}.top h1{{margin:0;font-size:20px}}.top p{{margin:4px 0 0;color:#b9c8d6}}.top a{{color:white;background:#2f6bff;padding:10px 14px;text-decoration:none;border-radius:4px}}main{{padding:24px;display:grid;grid-template-columns:repeat(auto-fit,minmax(410px,1fr));gap:24px}}.card{{background:white;border:1px solid #d7e0e7;border-radius:8px;padding:14px;box-shadow:0 3px 12px #071b2710}}header{{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px}}button{{border:1px solid #2f6bff;background:white;color:#2f6bff;padding:7px 10px;border-radius:4px;cursor:pointer}}.viewport{{width:384px;height:480px;margin:auto;overflow:hidden;background:#ddd;box-shadow:0 2px 10px #071b2730}}.slide{{position:relative;width:768px;height:960px;transform:scale(.5);transform-origin:top left;overflow:hidden}}.obj{{position:absolute;margin:0;padding:0}}.line{{overflow:visible}}.text div{{white-space:normal}}.reference img{{display:block;width:384px;height:480px}}.hidden{{display:none}}@media(max-width:500px){{main{{padding:10px;grid-template-columns:1fr}}.card{{padding:8px}}}}
</style></head><body><div class="top"><div><h1>USD Cash Allocation — Local Python Builder</h1><p>{len(slide_rows)} slides · {sum(len(v) for v in objects_by_slide.values())} objects · spreadsheet is the source of truth</p></div><a href="../config/presentation_config.xlsx">Open spreadsheet</a></div><main>{''.join(slide_cards)}</main>
<script>function toggleView(button){{const card=button.closest('.card');const model=card.querySelector('.model');const ref=card.querySelector('.reference');model.classList.toggle('hidden');ref.classList.toggle('hidden');button.textContent=model.classList.contains('hidden')?'Show model':'Show reference';}}</script>
</body></html>"""
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
    PREVIEW_FILE.write_text(page, encoding="utf-8")
    return PREVIEW_FILE


def main() -> None:
    parser = argparse.ArgumentParser(description="Create a browser preview from the presentation spreadsheet.")
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    parser.add_argument("--open", action="store_true")
    args = parser.parse_args()
    output = generate_preview(args.config)
    print(output)
    if args.open:
        webbrowser.open(output.as_uri())


if __name__ == "__main__":
    main()
