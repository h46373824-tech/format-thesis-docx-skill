#!/usr/bin/env python3
"""Extract a formatting-oriented profile from a DOCX or DOTX file.

The script uses only the Python standard library. It does not modify the input.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from typing import Any, Iterable, Optional
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def qn(prefix: str, local: str) -> str:
    return f"{{{NS[prefix]}}}{local}"


def attr(element: Optional[ET.Element], name: str, prefix: str = "w") -> Optional[str]:
    if element is None:
        return None
    return element.get(qn(prefix, name))


def read_xml(archive: zipfile.ZipFile, member: str) -> Optional[ET.Element]:
    try:
        return ET.fromstring(archive.read(member))
    except KeyError:
        return None
    except ET.ParseError as exc:
        raise ValueError(f"Invalid XML in {member}: {exc}") from exc


def paragraph_text(paragraph: ET.Element) -> str:
    pieces: list[str] = []
    for node in paragraph.iter():
        if node.tag == qn("w", "t"):
            pieces.append(node.text or "")
        elif node.tag == qn("w", "tab"):
            pieces.append("\t")
        elif node.tag in {qn("w", "br"), qn("w", "cr")}:
            pieces.append("\n")
    return "".join(pieces).strip()


def table_text(table: ET.Element) -> list[list[str]]:
    rows: list[list[str]] = []
    for row in table.findall("./w:tr", NS):
        values: list[str] = []
        for cell in row.findall("./w:tc", NS):
            parts = [
                paragraph_text(p)
                for p in cell.findall(".//w:p", NS)
                if paragraph_text(p)
            ]
            values.append(" | ".join(parts))
        rows.append(values)
    return rows


def twips_value(raw: Optional[str]) -> dict[str, Any] | None:
    if raw is None:
        return None
    try:
        twips = int(raw)
    except ValueError:
        return {"raw": raw}
    return {"twips": twips, "cm": round(twips / 1440 * 2.54, 3)}


def half_points(raw: Optional[str]) -> dict[str, Any] | None:
    if raw is None:
        return None
    try:
        value = int(raw)
    except ValueError:
        return {"raw": raw}
    return {"half_points": value, "pt": value / 2}


def on_off(element: Optional[ET.Element]) -> bool | None:
    if element is None:
        return None
    raw = attr(element, "val")
    if raw is None:
        return True
    return raw.lower() not in {"0", "false", "off", "no"}


def bool_value(raw: Optional[str]) -> bool | None:
    if raw is None:
        return None
    return raw.lower() not in {"0", "false", "off", "no"}


def parse_relationships(root: Optional[ET.Element]) -> dict[str, str]:
    if root is None:
        return {}
    relationships: dict[str, str] = {}
    for rel in root:
        rel_id = rel.get("Id")
        target = rel.get("Target")
        if rel_id and target:
            relationships[rel_id] = target
    return relationships


def style_record(style: ET.Element) -> dict[str, Any]:
    ppr = style.find("./w:pPr", NS)
    rpr = style.find("./w:rPr", NS)
    fonts = rpr.find("./w:rFonts", NS) if rpr is not None else None
    spacing = ppr.find("./w:spacing", NS) if ppr is not None else None
    indent = ppr.find("./w:ind", NS) if ppr is not None else None
    color = rpr.find("./w:color", NS) if rpr is not None else None
    justification = ppr.find("./w:jc", NS) if ppr is not None else None
    outline = ppr.find("./w:outlineLvl", NS) if ppr is not None else None

    return {
        "style_id": attr(style, "styleId"),
        "type": attr(style, "type"),
        "name": attr(style.find("./w:name", NS), "val"),
        "based_on": attr(style.find("./w:basedOn", NS), "val"),
        "next": attr(style.find("./w:next", NS), "val"),
        "default": bool_value(attr(style, "default")),
        "outline_level": attr(outline, "val"),
        "font": {
            "ascii": attr(fonts, "ascii"),
            "hansi": attr(fonts, "hAnsi"),
            "east_asia": attr(fonts, "eastAsia"),
            "cs": attr(fonts, "cs"),
        },
        "size": half_points(attr(rpr.find("./w:sz", NS), "val") if rpr is not None else None),
        "size_cs": half_points(
            attr(rpr.find("./w:szCs", NS), "val") if rpr is not None else None
        ),
        "bold": on_off(rpr.find("./w:b", NS)) if rpr is not None else None,
        "italic": on_off(rpr.find("./w:i", NS)) if rpr is not None else None,
        "color": attr(color, "val"),
        "alignment": attr(justification, "val"),
        "spacing": {
            "before": twips_value(attr(spacing, "before")),
            "after": twips_value(attr(spacing, "after")),
            "line": attr(spacing, "line"),
            "line_rule": attr(spacing, "lineRule"),
        },
        "indent": {
            "left": twips_value(attr(indent, "left")),
            "right": twips_value(attr(indent, "right")),
            "first_line": twips_value(attr(indent, "firstLine")),
            "hanging": twips_value(attr(indent, "hanging")),
        },
        "keep_next": on_off(ppr.find("./w:keepNext", NS)) if ppr is not None else None,
        "page_break_before": (
            on_off(ppr.find("./w:pageBreakBefore", NS)) if ppr is not None else None
        ),
    }


def heading_level_from_style(style: ET.Element) -> Optional[int]:
    style_id = (attr(style, "styleId") or "").replace(" ", "").lower()
    style_name = (
        attr(style.find("./w:name", NS), "val") or ""
    ).replace(" ", "").lower()
    for value in (style_id, style_name):
        match = re.search(r"(?:heading|标题)([1-9])", value)
        if match:
            return int(match.group(1))

    outline = style.find("./w:pPr/w:outlineLvl", NS)
    raw_outline = attr(outline, "val")
    if raw_outline is not None:
        try:
            return int(raw_outline) + 1
        except ValueError:
            return None
    return None


def build_style_heading_map(root: Optional[ET.Element]) -> dict[str, int]:
    if root is None:
        return {}
    result: dict[str, int] = {}
    for style in root.findall("./w:style", NS):
        style_id = attr(style, "styleId")
        level = heading_level_from_style(style)
        if style_id and level is not None:
            result[style_id] = level
    return result


def parse_styles(root: Optional[ET.Element]) -> dict[str, Any]:
    if root is None:
        return {
            "defaults": {},
            "selected": [],
            "outline_styles": [],
            "heading_style_map": {},
        }

    defaults: dict[str, Any] = {}
    doc_defaults = root.find("./w:docDefaults", NS)
    if doc_defaults is not None:
        rpr = doc_defaults.find("./w:rPrDefault/w:rPr", NS)
        ppr = doc_defaults.find("./w:pPrDefault/w:pPr", NS)
        fonts = rpr.find("./w:rFonts", NS) if rpr is not None else None
        spacing = ppr.find("./w:spacing", NS) if ppr is not None else None
        defaults = {
            "font": {
                "ascii": attr(fonts, "ascii"),
                "hansi": attr(fonts, "hAnsi"),
                "east_asia": attr(fonts, "eastAsia"),
                "cs": attr(fonts, "cs"),
            },
            "size": half_points(
                attr(rpr.find("./w:sz", NS), "val") if rpr is not None else None
            ),
            "spacing": {
                "before": twips_value(attr(spacing, "before")),
                "after": twips_value(attr(spacing, "after")),
                "line": attr(spacing, "line"),
                "line_rule": attr(spacing, "lineRule"),
            },
        }

    selected: list[dict[str, Any]] = []
    outline_styles: list[dict[str, Any]] = []
    wanted = {
        "normal",
        "title",
        "subtitle",
        "heading1",
        "heading2",
        "heading3",
        "toc1",
        "toc2",
        "toc3",
        "caption",
        "footnotetext",
        "bibliography",
    }

    for style in root.findall("./w:style", NS):
        record = style_record(style)
        style_id = (record.get("style_id") or "").replace(" ", "").lower()
        style_name = (record.get("name") or "").replace(" ", "").lower()
        if style_id in wanted or style_name in wanted:
            selected.append(record)
        if record.get("outline_level") is not None:
            outline_styles.append(record)

    return {
        "defaults": defaults,
        "selected": selected,
        "outline_styles": outline_styles,
        "heading_style_map": build_style_heading_map(root),
    }


def paragraph_record(
    paragraph: ET.Element, index: int, style_heading_map: dict[str, int]
) -> dict[str, Any]:
    ppr = paragraph.find("./w:pPr", NS)
    pstyle = ppr.find("./w:pStyle", NS) if ppr is not None else None
    outline = ppr.find("./w:outlineLvl", NS) if ppr is not None else None
    numpr = ppr.find("./w:numPr", NS) if ppr is not None else None
    num_id = numpr.find("./w:numId", NS) if numpr is not None else None
    ilvl = numpr.find("./w:ilvl", NS) if numpr is not None else None
    text = paragraph_text(paragraph)
    style_id = attr(pstyle, "val")

    heading_level: Optional[int] = style_heading_map.get(style_id or "")
    if heading_level is None and attr(outline, "val") is not None:
        try:
            heading_level = int(attr(outline, "val") or "") + 1
        except ValueError:
            heading_level = None

    heading_pattern = None
    patterns = [
        ("chapter_cn", r"^第[一二三四五六七八九十百0-9]+章(?:\s|[：:、.]|$)"),
        ("decimal", r"^\d+(?:\.\d+){0,4}(?:\s|[、.]|$)"),
        ("chinese_list", r"^[一二三四五六七八九十]+、"),
    ]
    for name, pattern in patterns:
        if re.search(pattern, text):
            heading_pattern = name
            break

    has_page_break = any(
        attr(br, "type") == "page" for br in paragraph.findall(".//w:br", NS)
    )

    return {
        "index": index,
        "text": text,
        "style_id": style_id,
        "outline_level": attr(outline, "val"),
        "heading_level": heading_level,
        "heading_pattern": heading_pattern,
        "numbering": {
            "num_id": attr(num_id, "val"),
            "level": attr(ilvl, "val"),
        },
        "page_break": has_page_break,
        "section_break": ppr is not None and ppr.find("./w:sectPr", NS) is not None,
    }


def section_record(section: ET.Element, index: int) -> dict[str, Any]:
    size = section.find("./w:pgSz", NS)
    margins = section.find("./w:pgMar", NS)
    number = section.find("./w:pgNumType", NS)
    section_type = section.find("./w:type", NS)
    columns = section.find("./w:cols", NS)
    title_page = section.find("./w:titlePg", NS)

    refs = []
    for ref in list(section.findall("./w:headerReference", NS)) + list(
        section.findall("./w:footerReference", NS)
    ):
        refs.append(
            {
                "kind": "header" if ref.tag == qn("w", "headerReference") else "footer",
                "type": attr(ref, "type"),
                "relationship_id": attr(ref, "id", "r"),
            }
        )

    return {
        "index": index,
        "type": attr(section_type, "val") or "nextPage",
        "page_size": {
            "width": twips_value(attr(size, "w")),
            "height": twips_value(attr(size, "h")),
            "orientation": attr(size, "orient") or "portrait",
        },
        "margins": {
            "top": twips_value(attr(margins, "top")),
            "right": twips_value(attr(margins, "right")),
            "bottom": twips_value(attr(margins, "bottom")),
            "left": twips_value(attr(margins, "left")),
            "header": twips_value(attr(margins, "header")),
            "footer": twips_value(attr(margins, "footer")),
            "gutter": twips_value(attr(margins, "gutter")),
        },
        "page_numbering": {
            "format": attr(number, "fmt"),
            "start": attr(number, "start"),
        },
        "different_first_page": on_off(title_page),
        "columns": {
            "count": attr(columns, "num") or "1",
            "spacing": twips_value(attr(columns, "space")),
        },
        "references": refs,
    }


def header_footer_profiles(
    archive: zipfile.ZipFile, relationships: dict[str, str]
) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for rel_id, target in relationships.items():
        target_path = target.replace("\\", "/")
        normalized = (
            posixpath.normpath(target_path.lstrip("/"))
            if target_path.startswith("/")
            else posixpath.normpath(posixpath.join("word", target_path))
        )
        basename = posixpath.basename(normalized)
        if not (basename.startswith("header") or basename.startswith("footer")):
            continue
        member = normalized
        root = read_xml(archive, member)
        if root is None:
            continue
        text = " ".join(
            part
            for part in (paragraph_text(p) for p in root.findall(".//w:p", NS))
            if part
        )
        instruction_text = " ".join(
            (node.text or "").strip()
            for node in root.findall(".//w:instrText", NS)
            if (node.text or "").strip()
        )
        records.append(
            {
                "relationship_id": rel_id,
                "part": member,
                "kind": "header" if basename.startswith("header") else "footer",
                "text": text,
                "fields": instruction_text,
            }
        )
    return records


def cover_candidates(
    body_children: Iterable[ET.Element],
    paragraphs: list[dict[str, Any]],
    limit: int = 80,
) -> dict[str, Any]:
    first_heading = next(
        (p["index"] for p in paragraphs if p.get("heading_level") and p.get("text")),
        None,
    )
    first_page_break = next(
        (p["index"] for p in paragraphs if p.get("page_break")),
        None,
    )
    cutoffs = [value for value in (first_heading, first_page_break) if value is not None]
    cutoff = min(cutoffs) if cutoffs else min(len(paragraphs), limit)

    selected_paragraphs = [
        p for p in paragraphs if p["index"] <= cutoff and p.get("text")
    ][:limit]

    tables: list[list[list[str]]] = []
    paragraph_index = -1
    for child in body_children:
        if child.tag == qn("w", "p"):
            paragraph_index += 1
            if paragraph_index > cutoff:
                break
        elif child.tag == qn("w", "tbl") and paragraph_index <= cutoff:
            rows = table_text(child)
            if any(any(cell.strip() for cell in row) for row in rows):
                tables.append(rows)

    return {
        "cutoff_paragraph_index": cutoff,
        "paragraphs": selected_paragraphs,
        "tables": tables,
    }


def collect_fields(root: ET.Element) -> list[str]:
    values = []
    for node in root.findall(".//w:instrText", NS):
        value = " ".join((node.text or "").split())
        if value:
            values.append(value)
    return values


def settings_record(root: Optional[ET.Element]) -> dict[str, bool]:
    """Return settings that materially affect field refresh and pagination."""
    if root is None:
        return {
            "update_fields_on_open": False,
            "even_and_odd_headers": False,
        }
    return {
        "update_fields_on_open": on_off(root.find("./w:updateFields", NS)) is True,
        "even_and_odd_headers": on_off(
            root.find("./w:evenAndOddHeaders", NS)
        ) is True,
    }


def pagination_warnings(
    settings: dict[str, bool], sections: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    """Flag section parity combinations that can create Word-only blank pages."""
    if not settings["even_and_odd_headers"]:
        return []

    warnings: list[dict[str, Any]] = []
    for section in sections[1:]:
        raw_start = section["page_numbering"]["start"]
        try:
            start = int(raw_start) if raw_start is not None else None
        except ValueError:
            start = None
        if (
            start is not None
            and start % 2 == 1
            and section["type"] in {"nextPage", "oddPage"}
        ):
            warnings.append(
                {
                    "code": "possible_virtual_even_page",
                    "section_index": section["index"],
                    "message": (
                        "Even/odd headers are enabled and this section restarts on an "
                        "odd page number. Word may insert a virtual blank even page at "
                        "the section boundary; verify the physical page count in Word."
                    ),
                }
            )
    return warnings


def profile(path: Path) -> dict[str, Any]:
    if path.suffix.lower() not in {".docx", ".dotx", ".docm", ".dotm"}:
        raise ValueError("Input must be a DOCX, DOTX, DOCM, or DOTM file")
    if not path.exists():
        raise FileNotFoundError(path)

    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    with zipfile.ZipFile(path) as archive:
        document = read_xml(archive, "word/document.xml")
        if document is None:
            raise ValueError("word/document.xml is missing")
        body = document.find("./w:body", NS)
        if body is None:
            raise ValueError("Document body is missing")

        styles_root = read_xml(archive, "word/styles.xml")
        settings = settings_record(read_xml(archive, "word/settings.xml"))
        styles = parse_styles(styles_root)
        style_heading_map = build_style_heading_map(styles_root)
        rels = parse_relationships(
            read_xml(archive, "word/_rels/document.xml.rels")
        )

        body_children = list(body)
        paragraphs = [
            paragraph_record(child, index, style_heading_map)
            for index, child in enumerate(
                child for child in body_children if child.tag == qn("w", "p")
            )
        ]
        sections = [
            section_record(section, index)
            for index, section in enumerate(document.findall(".//w:sectPr", NS))
        ]
        fields = collect_fields(document)
        warnings = pagination_warnings(settings, sections)
        names = set(archive.namelist())

        return {
            "schema_version": 1,
            "file": {
                "path": str(path.resolve()),
                "name": path.name,
                "size_bytes": path.stat().st_size,
                "sha256": digest,
                "macro_enabled": path.suffix.lower() in {".docm", ".dotm"},
            },
            "counts": {
                "paragraphs": len(paragraphs),
                "nonempty_paragraphs": sum(bool(p["text"]) for p in paragraphs),
                "tables": len(body.findall("./w:tbl", NS)),
                "images": len(
                    [
                        name
                        for name in names
                        if name.startswith("word/media/") and not name.endswith("/")
                    ]
                ),
                "sections": len(sections),
                "comments": int("word/comments.xml" in names),
                "footnotes": int("word/footnotes.xml" in names),
                "endnotes": int("word/endnotes.xml" in names),
            },
            "styles": styles,
            "settings": settings,
            "sections": sections,
            "warnings": warnings,
            "headers_footers": header_footer_profiles(archive, rels),
            "fields": fields,
            "has_toc_field": any(
                re.search(r"(?:^|\s)TOC(?:\s|$)", field, re.IGNORECASE)
                for field in fields
            ),
            "paragraphs": paragraphs,
            "cover_candidates": cover_candidates(body_children, paragraphs),
        }


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract a formatting profile from a Word OOXML file."
    )
    parser.add_argument("input", type=Path, help="Input .docx/.dotx/.docm/.dotm")
    parser.add_argument("--out", type=Path, help="Write JSON to this path")
    parser.add_argument(
        "--pretty",
        action="store_true",
        help="Pretty-print JSON to stdout when --out is omitted",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    try:
        result = profile(args.input)
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    indent = 2 if args.pretty or args.out else None
    payload = json.dumps(result, ensure_ascii=False, indent=indent)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(payload + "\n", encoding="utf-8")
        print(args.out)
    else:
        print(payload)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
