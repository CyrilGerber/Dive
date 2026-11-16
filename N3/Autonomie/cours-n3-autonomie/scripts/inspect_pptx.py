#!/usr/bin/env python3
"""List text-shape metadata to support manual extraction verification."""

from __future__ import annotations

import argparse
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("slides", nargs="*", type=int)
    args = parser.parse_args()
    with zipfile.ZipFile(args.source) as archive:
        selected = args.slides or range(1, 28)
        for number in selected:
            root = ET.fromstring(archive.read(f"ppt/slides/slide{number}.xml"))
            print(f"SLIDE {number}")
            for shape in root.findall(".//p:sp", NS):
                texts = [node.text or "" for node in shape.findall(".//a:t", NS)]
                if not any(text.strip() for text in texts):
                    continue
                offset = shape.find("p:spPr/a:xfrm/a:off", NS)
                ph = shape.find(".//p:ph", NS)
                sizes = [int(node.get("sz")) for node in shape.findall(".//a:rPr", NS) if node.get("sz")]
                name_node = shape.find(".//p:cNvPr", NS)
                ph_type = (ph.get("type") or "-") if ph is not None else "-"
                print(
                    "  ",
                    f"y={offset.get('y') if offset is not None else '-':>8}",
                    f"x={offset.get('x') if offset is not None else '-':>8}",
                    f"size={max(sizes) if sizes else '-':>4}",
                    f"ph={ph_type:>8}",
                    f"name={name_node.get('name') if name_node is not None else '-'}",
                    repr(" | ".join(texts)[:180]),
                )
                for paragraph in shape.findall("p:txBody/a:p", NS):
                    paragraph_text = "".join(node.text or "" for node in paragraph.findall(".//a:t", NS))
                    paragraph_sizes = [
                        int(node.get("sz")) for node in paragraph.iter() if node.get("sz", "").isdigit()
                    ]
                    if paragraph_text.strip():
                        print(f"      paragraph size={max(paragraph_sizes, default=0):>4} {paragraph_text!r}")


if __name__ == "__main__":
    main()
