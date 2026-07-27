#!/usr/bin/env python3
"""Extract the SOFR local-model package stored in the payload folder."""
from __future__ import annotations

import base64
import io
import zipfile
from pathlib import Path


def main() -> None:
    root = Path(__file__).resolve().parent
    parts = sorted((root / "payload").glob("part*.txt"))
    if not parts:
        raise FileNotFoundError("No payload files were found.")
    encoded = "".join(part.read_text(encoding="ascii").strip() for part in parts)
    archive = base64.b85decode(encoded)
    with zipfile.ZipFile(io.BytesIO(archive), "r") as package:
        package.extractall(root)
    print(f"SOFR local model extracted to: {root}")
    print("Next commands:")
    print("  python -m pip install -r requirements_sofr_local.txt")
    print("  python sofr_local_model.py --create-template SOFR_Local_Model.xlsx")


if __name__ == "__main__":
    main()
