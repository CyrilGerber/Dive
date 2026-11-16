#!/usr/bin/env python3
"""Extract faithful text and packaged assets from a PowerPoint Open XML file.

This script intentionally uses only Python's standard library. It never writes to
the source PPTX. Generated Markdown is a transcription layer, not a replacement
for the immutable source presentation.
"""

from __future__ import annotations

import argparse
import hashlib
import posixpath
import re
import shutil
import sys
import unicodedata
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R_EMBED = f"{{{REL_NS}}}embed"
R_LINK = f"{{{REL_NS}}}link"


@dataclass
class Relationship:
    rel_id: str
    rel_type: str
    target: str
    external: bool = False


@dataclass
class Visual:
    kind: str
    shape_name: str
    relationship_id: str = ""
    asset: str = ""
    detail: str = ""


@dataclass
class SlideResult:
    number: int
    title: str = ""
    blocks: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    visuals: list[Visual] = field(default_factory=list)
    observations: list[str] = field(default_factory=list)
    native_shapes: int = 0
    output_name: str = ""
    status: str = "extracted"


def qname(prefix: str, local: str) -> str:
    return f"{{{NS[prefix]}}}{local}"


def read_xml(archive: zipfile.ZipFile, name: str) -> ET.Element:
    return ET.fromstring(archive.read(name))


def rels_name(part_name: str) -> str:
    parent = posixpath.dirname(part_name)
    base = posixpath.basename(part_name)
    return posixpath.join(parent, "_rels", base + ".rels")


def normalized_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def relationships(archive: zipfile.ZipFile, source_part: str) -> dict[str, Relationship]:
    rel_part = rels_name(source_part)
    if rel_part not in archive.namelist():
        return {}
    root = read_xml(archive, rel_part)
    result: dict[str, Relationship] = {}
    for node in root.findall("pr:Relationship", NS):
        rel_id = node.get("Id", "")
        result[rel_id] = Relationship(
            rel_id=rel_id,
            rel_type=node.get("Type", "").rsplit("/", 1)[-1],
            target=normalized_target(source_part, node.get("Target", "")),
            external=node.get("TargetMode") == "External",
        )
    return result


def slide_parts_in_order(archive: zipfile.ZipFile) -> list[str]:
    presentation = "ppt/presentation.xml"
    root = read_xml(archive, presentation)
    rels = relationships(archive, presentation)
    parts: list[str] = []
    for slide_id in root.findall("p:sldIdLst/p:sldId", NS):
        rel_id = slide_id.get(qname("r", "id"), "")
        if rel_id in rels:
            parts.append(rels[rel_id].target)
    return parts


def paragraph_text(paragraph: ET.Element) -> str:
    chunks: list[str] = []
    for node in paragraph.iter():
        if node.tag == qname("a", "t"):
            chunks.append(node.text or "")
        elif node.tag == qname("a", "br"):
            chunks.append("\n")
        elif node.tag == qname("a", "tab"):
            chunks.append("\t")
    return "".join(chunks)


def text_body_lines(tx_body: ET.Element) -> list[str]:
    lines: list[str] = []
    for paragraph in tx_body.findall("a:p", NS):
        text = paragraph_text(paragraph)
        ppr = paragraph.find("a:pPr", NS)
        level = int(ppr.get("lvl", "0")) if ppr is not None else 0
        is_bullet = False
        if ppr is not None:
            is_bullet = ppr.find("a:buChar", NS) is not None or ppr.find("a:buAutoNum", NS) is not None
        if text:
            prefix = ("  " * level + "- ") if is_bullet else ""
            lines.append(prefix + text)
        else:
            lines.append("")
    while lines and not lines[-1]:
        lines.pop()
    return lines


def paragraph_font_size(paragraph: ET.Element) -> int:
    sizes: list[int] = []
    for node in paragraph.iter():
        value = node.get("sz")
        if value and value.isdigit():
            sizes.append(int(value))
    return max(sizes, default=0)


def shape_position(shape: ET.Element) -> tuple[int, int]:
    offset = shape.find("p:spPr/a:xfrm/a:off", NS)
    if offset is None:
        return (10**9, 10**9)
    return (int(offset.get("x", "0")), int(offset.get("y", "0")))


def shape_name(shape: ET.Element) -> str:
    node = shape.find(".//p:cNvPr", NS)
    return node.get("name", "") if node is not None else ""


def placeholder_type(shape: ET.Element) -> str:
    node = shape.find(".//p:nvPr/p:ph", NS)
    return node.get("type", "") if node is not None else ""


def yaml_quote(value: str) -> str:
    escaped = value.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")
    return f'"{escaped}"'


def markdown_escape_heading(value: str) -> str:
    return value.replace("\\", "\\\\").replace("#", "\\#")


def slugify(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_value = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    slug = re.sub(r"[^a-zA-Z0-9]+", "-", ascii_value).strip("-").lower()
    return slug[:55] or "slide"


def asset_path_for_package_part(target: str) -> str:
    path = PurePosixPath(target)
    if target.startswith("ppt/media/"):
        return f"assets/original/{path.name}"
    if target.startswith("ppt/charts/"):
        return f"assets/original/charts/{path.name}"
    if target.startswith("ppt/diagrams/"):
        return f"assets/original/diagrams/{path.name}"
    if target.startswith("ppt/embeddings/"):
        return f"assets/original/embeddings/{path.name}"
    return ""


def visual_from_rel(kind: str, name: str, rel_id: str, rels: dict[str, Relationship]) -> Visual:
    rel = rels.get(rel_id)
    if rel is None:
        return Visual(kind, name, rel_id, detail="Relation absente du paquet")
    if rel.external:
        return Visual(kind, name, rel_id, detail=f"Cible externe : {rel.target}")
    return Visual(kind, name, rel_id, asset=asset_path_for_package_part(rel.target), detail=rel.rel_type)


def extract_table(frame: ET.Element) -> str:
    table = frame.find(".//a:tbl", NS)
    if table is None:
        return ""
    rows: list[list[str]] = []
    for tr in table.findall("a:tr", NS):
        row: list[str] = []
        for tc in tr.findall("a:tc", NS):
            cell_lines: list[str] = []
            for tx_body in tc.findall("a:txBody", NS):
                cell_lines.extend(text_body_lines(tx_body))
            row.append("<br>".join(cell_lines).replace("|", "\\|"))
        rows.append(row)
    if not rows:
        return ""
    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]
    output = ["| " + " | ".join(rows[0]) + " |", "| " + " | ".join(["---"] * width) + " |"]
    output.extend("| " + " | ".join(row) + " |" for row in rows[1:])
    return "\n".join(output)


def parse_slide(archive: zipfile.ZipFile, slide_part: str, number: int) -> SlideResult:
    root = read_xml(archive, slide_part)
    rels = relationships(archive, slide_part)
    result = SlideResult(number=number)
    text_shapes: list[dict[str, object]] = []

    shape_tree = root.find("p:cSld/p:spTree", NS)
    if shape_tree is None:
        result.observations.append("Arbre de formes introuvable dans la diapositive.")
        result.status = "needs_verification"
        return result

    def visit(container: ET.Element) -> None:
        for child in list(container):
            local = child.tag.rsplit("}", 1)[-1]
            if local in {"nvGrpSpPr", "grpSpPr"}:
                continue
            name = shape_name(child)
            if local == "sp":
                tx_body = child.find("p:txBody", NS)
                lines = text_body_lines(tx_body) if tx_body is not None else []
                ph_type = placeholder_type(child)
                if lines and any(line.strip() for line in lines):
                    paragraphs = tx_body.findall("a:p", NS) if tx_body is not None else []
                    sizes = [paragraph_font_size(paragraph) for paragraph in paragraphs]
                    x, y = shape_position(child)
                    text_shapes.append(
                        {"lines": lines, "sizes": sizes, "placeholder": ph_type, "x": x, "y": y}
                    )
                else:
                    geometry = child.find("p:spPr/a:prstGeom", NS)
                    custom = child.find("p:spPr/a:custGeom", NS)
                    if geometry is not None or custom is not None:
                        result.native_shapes += 1
            elif local == "pic":
                blip = child.find(".//a:blip", NS)
                rel_id = ""
                if blip is not None:
                    rel_id = blip.get(R_EMBED, "") or blip.get(R_LINK, "")
                result.visuals.append(visual_from_rel("image", name, rel_id, rels))
            elif local == "graphicFrame":
                table_md = extract_table(child)
                if table_md:
                    result.blocks.append(table_md)
                else:
                    graphic_data = child.find(".//a:graphicData", NS)
                    uri = graphic_data.get("uri", "") if graphic_data is not None else ""
                    rel_nodes = [n for n in child.iter() for key in (R_EMBED, R_LINK, qname("r", "id")) if n.get(key)]
                    rel_ids: list[str] = []
                    for node in rel_nodes:
                        for key in (R_EMBED, R_LINK, qname("r", "id")):
                            rid = node.get(key)
                            if rid and rid not in rel_ids:
                                rel_ids.append(rid)
                    kind = "smartart" if "diagram" in uri else "chart" if "chart" in uri else "graphic-frame"
                    if rel_ids:
                        result.visuals.extend(visual_from_rel(kind, name, rid, rels) for rid in rel_ids)
                    else:
                        result.visuals.append(Visual(kind, name, detail=uri or "Cadre graphique non identifié"))
            elif local == "grpSp":
                visit(child)
            elif local == "cxnSp":
                result.native_shapes += 1
            elif local == "contentPart":
                result.visuals.append(Visual("content-part", name, detail="Élément de contenu non extractible"))

    visit(shape_tree)
    selected_shape: dict[str, object] | None = None
    selected_indexes: list[int] = []
    placeholder_shapes = [shape for shape in text_shapes if shape["placeholder"] in {"title", "ctrTitle"}]
    if placeholder_shapes:
        selected_shape = placeholder_shapes[0]
        selected_indexes = [i for i, line in enumerate(selected_shape["lines"]) if str(line).strip()]
        if len(placeholder_shapes) > 1:
            result.observations.append(
                "Plusieurs espaces réservés de titre détectés ; seul le premier est utilisé comme titre Markdown."
            )
    else:
        # Many legacy decks use an ordinary text box as a visual title. Restrict
        # inference to large text near the top and away from side navigation.
        candidates = [
            shape
            for shape in text_shapes
            if int(shape["y"]) <= 1_100_000
            and 1_000_000 <= int(shape["x"]) <= 8_000_000
            and max(shape["sizes"], default=0) >= 3_000
        ]
        if candidates:
            selected_shape = min(candidates, key=lambda shape: (int(shape["y"]), int(shape["x"])))
            selected_indexes = [
                i
                for i, (line, size) in enumerate(zip(selected_shape["lines"], selected_shape["sizes"]))
                if str(line).strip() and int(size) >= 2_200
            ]
        else:
            central = [
                shape
                for shape in text_shapes
                if 1_000_000 <= int(shape["x"]) <= 9_000_000 and shape["placeholder"] != "sldNum"
            ]
            if central:
                fallback = max(central, key=lambda shape: (max(shape["sizes"], default=0), -int(shape["y"])))
                largest = max(fallback["sizes"], default=0)
                if largest >= 3_000 or len(central) == 1:
                    selected_shape = fallback
                    selected_indexes = [
                        i
                        for i, (line, size) in enumerate(zip(fallback["lines"], fallback["sizes"]))
                        if str(line).strip() and (int(size) == largest or len(central) == 1)
                    ]

    if selected_shape is not None and selected_indexes:
        result.title = "\n".join(
            str(selected_shape["lines"][index]).strip() for index in selected_indexes
        ).strip()

    for shape in text_shapes:
        lines = list(shape["lines"])
        if shape is selected_shape:
            lines = [line for index, line in enumerate(lines) if index not in selected_indexes]
        while lines and not str(lines[0]).strip():
            lines.pop(0)
        while lines and not str(lines[-1]).strip():
            lines.pop()
        if lines:
            result.blocks.append("\n".join(str(line) for line in lines))

    notes_rel = next((rel for rel in rels.values() if rel.rel_type == "notesSlide" and not rel.external), None)
    if notes_rel and notes_rel.target in archive.namelist():
        notes_root = read_xml(archive, notes_rel.target)
        body_notes: list[str] = []
        other_notes: list[str] = []
        for shape in notes_root.findall(".//p:sp", NS):
            tx_body = shape.find("p:txBody", NS)
            lines = text_body_lines(tx_body) if tx_body is not None else []
            if not lines:
                continue
            ph_type = placeholder_type(shape)
            if ph_type == "body":
                body_notes.extend(lines)
            elif ph_type not in {"hdr", "ftr", "dt", "sldNum"}:
                other_notes.extend(lines)
        result.notes = body_notes or other_notes

    if result.native_shapes:
        result.observations.append(
            f"{result.native_shapes} forme(s) vectorielle(s) PowerPoint native(s) sans texte : "
            "elles restent conservées dans le PPTX source mais ne constituent pas des fichiers image séparés."
        )
    unresolved = [v for v in result.visuals if not v.asset and v.kind not in {"image"}]
    missing_images = [v for v in result.visuals if v.kind == "image" and not v.asset]
    if unresolved:
        result.observations.append(
            f"{len(unresolved)} élément(s) graphique(s) complexe(s) à vérifier dans PowerPoint (graphique, SmartArt ou contenu incorporé)."
        )
    if missing_images:
        result.observations.append(f"{len(missing_images)} image(s) liée(s) ou relation(s) d'image à vérifier manuellement.")
    if result.observations:
        result.status = "needs_verification"
    return result


def write_markdown(project: Path, result: SlideResult) -> None:
    visuals_yaml = "visuals: []"
    if result.visuals:
        rows = ["visuals:"]
        for visual in result.visuals:
            rows.append(f"  - kind: {yaml_quote(visual.kind)}")
            rows.append(f"    source_shape: {yaml_quote(visual.shape_name)}")
            rows.append(f"    relationship_id: {yaml_quote(visual.relationship_id)}")
            rows.append(f"    asset: {yaml_quote('../' + visual.asset if visual.asset else '')}")
        visuals_yaml = "\n".join(rows)

    title_heading = markdown_escape_heading(result.title) if result.title else "<!-- Titre non extractible -->"
    content = "\n\n".join(result.blocks)
    notes = "\n".join(result.notes)
    visual_lines: list[str] = []
    for visual in result.visuals:
        label = visual.shape_name or visual.kind
        if visual.asset:
            visual_lines.append(f"- [{label}](../{visual.asset}) — type : `{visual.kind}`")
        else:
            suffix = f" — {visual.detail}" if visual.detail else ""
            visual_lines.append(f"- {label} — type : `{visual.kind}`{suffix}")
    observations = "\n".join(f"- {item}" for item in result.observations)
    visuals_markdown = "\n".join(visual_lines)

    md = (
        "---\n"
        f"id: slide-{result.number:03d}\n"
        f"source_slide: {result.number}\n"
        f"status: {result.status}\n"
        "layout: original\n"
        f"{visuals_yaml}\n"
        "---\n\n"
        f"# {title_heading}\n\n"
        "## Contenu\n\n"
        f"{content}\n\n"
        "## Notes présentateur\n\n"
        f"{notes}\n\n"
        "## Visuels\n\n"
        f"{visuals_markdown}\n\n"
        "## Observations techniques\n\n"
        f"{observations}\n"
    )
    (project / "slides" / result.output_name).write_text(md, encoding="utf-8", newline="\n")


def validated_file_for_slide(project: Path, slide_number: int) -> Path | None:
    """Return an existing validated Markdown file for a source slide, if any."""
    for candidate in (project / "slides").glob("*.md"):
        frontmatter = candidate.read_text(encoding="utf-8")[:4_000]
        source_match = re.search(r"^source_slide:\s*(\d+)\s*$", frontmatter, re.MULTILINE)
        status_match = re.search(r"^status:\s*([a-z_]+)\s*$", frontmatter, re.MULTILINE)
        if (
            source_match
            and int(source_match.group(1)) == slide_number
            and status_match
            and status_match.group(1) == "validated"
        ):
            return candidate
    return None


def copy_packaged_assets(archive: zipfile.ZipFile, project: Path) -> list[str]:
    prefixes = {
        "ppt/media/": project / "assets" / "original",
        "ppt/charts/": project / "assets" / "original" / "charts",
        "ppt/diagrams/": project / "assets" / "original" / "diagrams",
        "ppt/embeddings/": project / "assets" / "original" / "embeddings",
    }
    copied: list[str] = []
    for member in archive.namelist():
        if member.endswith("/"):
            continue
        for prefix, destination in prefixes.items():
            if member.startswith(prefix):
                destination.mkdir(parents=True, exist_ok=True)
                out = destination / PurePosixPath(member).name
                with archive.open(member) as source, out.open("wb") as target:
                    shutil.copyfileobj(source, target)
                copied.append(out.relative_to(project).as_posix())
                break
    return sorted(copied)


def write_manifest(project: Path, source_name: str, source_hash: str, results: list[SlideResult]) -> None:
    lines = [
        "version: 1",
        f"source: {yaml_quote('sources/' + source_name)}",
        f"source_sha256: {yaml_quote(source_hash)}",
        f"slide_count: {len(results)}",
        "slides:",
    ]
    for result in results:
        lines.extend(
            [
                f"  - id: slide-{result.number:03d}",
                f"    source_slide: {result.number}",
                f"    title: {yaml_quote(result.title)}",
                f"    status: {result.status}",
                f"    file: {yaml_quote('slides/' + result.output_name)}",
            ]
        )
    (project / "manifest.yml").write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")


def write_report(project: Path, source: Path, source_hash: str, results: list[SlideResult], assets: list[str]) -> None:
    status_counts = Counter(result.status for result in results)
    lines = [
        "# Rapport d’extraction",
        "",
        f"- Source : `{source.name}`",
        f"- SHA-256 : `{source_hash}`",
        f"- Diapositives référencées : **{len(results)}**",
        f"- Fichiers d’actifs extraits : **{len(assets)}**",
        f"- Statuts : " + ", ".join(f"`{key}` = {value}" for key, value in sorted(status_counts.items())),
        "",
        "## Contrôles",
        "",
        f"- {'OK' if len(results) == 27 else 'À vérifier'} — le nombre attendu est 27.",
        f"- {'OK' if len({r.number for r in results}) == len(results) else 'À vérifier'} — identifiants de diapositives uniques.",
        f"- {'OK' if all((project / 'slides' / r.output_name).is_file() for r in results) else 'À vérifier'} — chaque entrée du manifeste possède un fichier Markdown.",
        "- Le fichier source a uniquement été lu ; l’extraction est reproductible avec `scripts/extract_pptx.py`.",
        "",
        "## Éléments nécessitant une vérification manuelle",
        "",
    ]
    flagged = [result for result in results if result.observations]
    if not flagged:
        lines.append("Aucun élément technique douteux détecté automatiquement.")
    else:
        for result in flagged:
            title = result.title or "sans titre extractible"
            lines.append(f"### Diapositive {result.number} — {title}")
            lines.append("")
            lines.extend(f"- {observation}" for observation in result.observations)
            lines.append("")
    lines.extend(
        [
            "## Limites de l’extraction automatique",
            "",
            "- Les formes vectorielles PowerPoint natives sont conservées dans le PPTX source, mais ne sont pas converties en SVG.",
            "- Les graphiques, SmartArt, objets OLE et médias complexes sont recensés et leurs composants empaquetés sont copiés lorsqu’ils existent.",
            "- L’ordre du texte suit l’ordre des objets dans le XML de la diapositive ; une vérification visuelle reste nécessaire lorsque la mise en page est complexe.",
            "- Aucun texte pédagogique n’a été corrigé, reformulé ou complété.",
            "",
            "## Actifs extraits",
            "",
        ]
    )
    lines.extend(f"- `{asset}`" for asset in assets)
    (project / "EXTRACTION_REPORT.md").write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path, help="PPTX source path")
    parser.add_argument("--project", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--expected-slides", type=int, default=27)
    args = parser.parse_args()

    source = args.source.resolve()
    project = args.project.resolve()
    if not source.is_file():
        parser.error(f"Source introuvable : {source}")
    if source.parent == project / "slides" or source.parent == project / "dist":
        parser.error("La source doit rester dans le dossier immuable sources/.")

    source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    (project / "slides").mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as archive:
        slide_parts = slide_parts_in_order(archive)
        results = [parse_slide(archive, part, index) for index, part in enumerate(slide_parts, start=1)]
        assets = copy_packaged_assets(archive, project)

    used_names: set[str] = set()
    for result in results:
        validated = validated_file_for_slide(project, result.number)
        if validated is not None:
            result.output_name = validated.name
            result.status = "validated"
            used_names.add(validated.name)
            continue
        base = f"{result.number:03d}-{slugify(result.title)}"
        name = base + ".md"
        suffix = 2
        while name in used_names:
            name = f"{base}-{suffix}.md"
            suffix += 1
        used_names.add(name)
        result.output_name = name
        write_markdown(project, result)

    write_manifest(project, source.name, source_hash, results)
    write_report(project, source, source_hash, results, assets)
    print(f"Extraction terminée : {len(results)} diapositive(s), {len(assets)} actif(s).")
    if len(results) != args.expected_slides:
        print(f"ERREUR : {args.expected_slides} diapositives attendues.", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
