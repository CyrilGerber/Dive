#!/usr/bin/env python3
"""Validate the extracted course project without third-party dependencies."""

from __future__ import annotations

import hashlib
import re
import sys
from pathlib import Path


def main() -> int:
    project = Path(__file__).resolve().parents[1]
    manifest_path = project / "manifest.yml"
    manifest = manifest_path.read_text(encoding="utf-8")
    errors: list[str] = []

    count_match = re.search(r"^slide_count:\s*(\d+)\s*$", manifest, re.MULTILINE)
    expected_count = int(count_match.group(1)) if count_match else -1
    entries = re.findall(
        r"  - id: (slide-\d{3})\n"
        r"    source_slide: (\d+)\n"
        r"    title: .*?\n"
        r"    status: ([a-z_]+)\n"
        r'    file: "([^"]+)"',
        manifest,
    )
    if expected_count != 27:
        errors.append(f"slide_count vaut {expected_count}, 27 attendu")
    if len(entries) != expected_count:
        errors.append(f"{len(entries)} entrées de manifeste pour slide_count={expected_count}")
    if len({entry[0] for entry in entries}) != len(entries):
        errors.append("identifiants de diapositives non uniques")
    if len({entry[1] for entry in entries}) != len(entries):
        errors.append("numéros source_slide non uniques")

    allowed_statuses = {"extracted", "in_review", "validated", "needs_verification"}
    for slide_id, source_number, status, relative_file in entries:
        if status not in allowed_statuses:
            errors.append(f"{slide_id}: statut inconnu {status}")
        markdown_path = project / relative_file
        if not markdown_path.is_file():
            errors.append(f"{slide_id}: fichier absent {relative_file}")
            continue
        content = markdown_path.read_text(encoding="utf-8")
        if f"id: {slide_id}" not in content:
            errors.append(f"{slide_id}: identifiant incohérent dans le Markdown")
        if f"source_slide: {source_number}" not in content:
            errors.append(f"{slide_id}: source_slide incohérent dans le Markdown")
        for link in re.findall(r"\]\(\.\./(assets/original/[^)]+)\)", content):
            if not (project / link).is_file():
                errors.append(f"{slide_id}: actif lié absent {link}")

    source_match = re.search(r'^source:\s*"([^"]+)"$', manifest, re.MULTILINE)
    hash_match = re.search(r'^source_sha256:\s*"([0-9a-f]{64})"$', manifest, re.MULTILINE)
    if source_match and hash_match:
        source = project / source_match.group(1)
        if not source.is_file():
            errors.append(f"source absente : {source}")
        else:
            actual_hash = hashlib.sha256(source.read_bytes()).hexdigest()
            if actual_hash != hash_match.group(1):
                errors.append("le hash de la source ne correspond pas au manifeste")
    else:
        errors.append("source ou source_sha256 absent du manifeste")

    markdown_count = len(list((project / "slides").glob("*.md")))
    if markdown_count != expected_count:
        errors.append(f"{markdown_count} Markdown présents, {expected_count} attendus")

    if errors:
        print("VALIDATION ÉCHOUÉE")
        for error in errors:
            print(f"- {error}")
        return 1
    print(f"VALIDATION OK — {expected_count} diapositives, manifeste, source et actifs cohérents.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
