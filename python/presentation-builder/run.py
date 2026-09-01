from __future__ import annotations

import argparse
import os
import sys
import webbrowser
from pathlib import Path


ROOT = Path(__file__).resolve().parent


def _load_builder():
    try:
        from build_presentation import (
            ASSET_DIR,
            DEFAULT_CONFIG,
            build_presentation,
            configured_output_path,
            load_model,
        )
        from preview import generate_preview
    except ModuleNotFoundError as exc:
        package = exc.name or "a required Python package"
        raise RuntimeError(
            f"Missing dependency: {package}. Ask your corporate Python administrator "
            "to install the approved packages listed in requirements.txt."
        ) from exc
    return ASSET_DIR, DEFAULT_CONFIG, build_presentation, configured_output_path, load_model, generate_preview


def inspect_model() -> dict:
    ASSET_DIR, config_path, _, configured_output_path, load_model, _ = _load_builder()
    if not config_path.exists():
        raise FileNotFoundError(f"Configuration workbook not found: {config_path}")
    model = load_model(config_path)
    missing_assets: list[str] = []
    for row in model["objects"]:
        if not bool(row.get("Enabled")):
            continue
        if str(row.get("Object_Type") or "") not in {"Image", "ChartImage"}:
            continue
        asset_name = Path(str(row.get("Asset_ID") or "")).name
        if not asset_name or not (ASSET_DIR / asset_name).exists():
            missing_assets.append(asset_name or "<blank Asset_ID>")
    return {
        "model": model,
        "config": config_path,
        "output": configured_output_path(config_path),
        "missing_assets": sorted(set(missing_assets)),
    }


def print_status() -> bool:
    details = inspect_model()
    model = details["model"]
    enabled = sum(bool(row.get("Enabled")) for row in model["objects"])
    print("Presentation Builder status")
    print(f"  Configuration : {details['config']}")
    print(f"  Output        : {details['output']}")
    print(f"  Slides        : {len(model['slides'])}")
    print(f"  Objects       : {enabled}")
    print(f"  Text runs     : {len(model['runs'])}")
    print(f"  Formats       : {len(model['formats'])}")
    if details["missing_assets"]:
        print("  Missing assets:")
        for name in details["missing_assets"]:
            print(f"    - {name}")
        return False
    print("  Validation    : PASS")
    return True


def build() -> None:
    _, config_path, build_presentation, _, _, _ = _load_builder()
    if not print_status():
        raise RuntimeError("Build stopped because required assets are missing.")
    output = build_presentation(config_path)
    print(f"Presentation created: {output}")


def preview(open_browser: bool = False) -> None:
    _, config_path, _, _, _, generate_preview = _load_builder()
    if not print_status():
        raise RuntimeError("Preview stopped because required assets are missing.")
    output = generate_preview(config_path)
    print(f"Preview created: {output}")
    if open_browser:
        webbrowser.open(output.as_uri())


def open_config() -> None:
    _, config_path, _, _, _, _ = _load_builder()
    if not config_path.exists():
        raise FileNotFoundError(config_path)
    if os.name != "nt":
        raise RuntimeError("The open-config command is intended for Windows.")
    os.startfile(config_path)  # type: ignore[attr-defined]


def interactive_menu() -> None:
    actions = {
        "1": ("Validate and show status", print_status),
        "2": ("Build presentation", build),
        "3": ("Generate browser preview", preview),
        "4": ("Generate and open browser preview", lambda: preview(True)),
        "5": ("Open configuration workbook", open_config),
    }
    while True:
        print("\nUSD Cash Allocation — Presentation Builder")
        for key, (label, _) in actions.items():
            print(f"  {key}. {label}")
        print("  0. Exit")
        choice = input("Select an option: ").strip()
        if choice == "0":
            return
        action = actions.get(choice)
        if action is None:
            print("Invalid option.")
            continue
        try:
            action[1]()
        except Exception as exc:
            print(f"ERROR: {exc}", file=sys.stderr)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Corporate-friendly Python entry point for the spreadsheet-driven presentation builder."
    )
    subparsers = parser.add_subparsers(dest="command")
    subparsers.add_parser("status", help="Validate the workbook and list model counts.")
    subparsers.add_parser("validate", help="Validate the workbook and required assets.")
    subparsers.add_parser("build", help="Build the PowerPoint presentation.")
    preview_parser = subparsers.add_parser("preview", help="Generate the static browser preview.")
    preview_parser.add_argument("--open", action="store_true", help="Open the preview in the default browser.")
    subparsers.add_parser("open-config", help="Open the configuration workbook in Windows.")
    subparsers.add_parser("gui", help="Open the optional Tkinter manager.")
    args = parser.parse_args()

    try:
        if args.command is None:
            interactive_menu()
        elif args.command in {"status", "validate"}:
            if not print_status():
                raise RuntimeError("Validation failed.")
        elif args.command == "build":
            build()
        elif args.command == "preview":
            preview(args.open)
        elif args.command == "open-config":
            open_config()
        elif args.command == "gui":
            from app import Manager

            Manager().mainloop()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()

