#!/usr/bin/env python3
"""Extract a formatting-oriented profile from a DOCX or DOTX file.

The script uses only the Python standard library. It does not modify the input.
Resolved style values include basedOn inheritance and theme fonts. Paragraph
records include direct formatting and text from content controls, table cells,
and OMML. Cover candidates stop at the first heading, page break, or section
break. This profile does not prove desktop Word pagination.
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
from typing import Any, Optional
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}

MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

MAX_PACKAGE_BYTES = 64 * 1024 * 1024
MAX_UNCOMPRESSED_BYTES = 256 * 1024 * 1024
MAX_XML_BYTES = 32 * 1024 * 1024
MAX_PARAGRAPHS = 20000
MAX_COVER_PARAGRAPHS = 80
MAX_COVER_TABLES = 20
MAX_SELECTED_STYLES = 200

FONT_KEYS = ("ascii", "hansi", "east_asia", "cs")
EXPLICIT_FONT_ATTRS = {
    "ascii": "ascii",
    "hansi": "hAnsi",
    "east_asia": "eastAsia",
    "cs": "cs",
}
THEME_FONT_ATTRS = {
    "ascii": "asciiTheme",
    "hansi": "hAnsiTheme",
    "east_asia": "eastAsiaTheme",
    "cs": "cstheme",
}
NOTE_SKIP_TYPES = {"separator", "continuationSeparator", "continuationNotice"}
HEADING_NAME_RE = re.compile(r"(?:heading|标题)([1-9])(?!\d)", re.IGNORECASE)
DTD_RE = re.compile(br"<!DOCTYPE|<!ENTITY", re.IGNORECASE)


def qn(prefix: str, local: str) -> str:
    return f"{{{NS[prefix]}}}{local}"


def mc(local: str) -> str:
    return f"{{{MC_NS}}}{local}"


def attr(element: Optional[ET.Element], name: str, prefix: str = "w") -> Optional[str]:
    if element is None:
        return None
    return element.get(qn(prefix, name))


def selected_child_elements(element: ET.Element) -> list[ET.Element]:
    """Return children, keeping only one branch of each AlternateContent.

    Fallback is preferred because it is the branch every consumer can read.
    The other branch is ignored so Choice and Fallback are not both counted.
    """
    children: list[ET.Element] = []
    for child in list(element):
        if child.tag != mc("AlternateContent"):
            children.append(child)
            continue
        fallback = next((node for node in list(child) if node.tag == mc("Fallback")), None)
        choice = next((node for node in list(child) if node.tag == mc("Choice")), None)
        chosen = fallback if fallback is not None else choice
        if chosen is not None:
            children.extend(list(chosen))
    return children


def iter_selected(element: ET.Element):
    yield element
    for child in selected_child_elements(element):
        yield from iter_selected(child)


def first_selected(element: Optional[ET.Element], local: str) -> Optional[ET.Element]:
    if element is None:
        return None
    wanted = qn("w", local)
    for child in selected_child_elements(element):
        if child.tag == wanted:
            return child
    return None


def read_capped(archive: zipfile.ZipFile, member: str) -> bytes:
    info = archive.getinfo(member)
    if info.file_size > MAX_XML_BYTES:
        raise ValueError(
            f"{member} declares uncompressed size {info.file_size}, "
            f"limit is {MAX_XML_BYTES}"
        )
    payload = archive.read(member)
    if len(payload) > MAX_XML_BYTES:
        raise ValueError(f"{member} exceeds uncompressed XML limit {MAX_XML_BYTES}")
    if DTD_RE.search(payload):
        raise ValueError(f"DTD or entity declarations are not allowed in {member}")
    return payload


def read_xml(archive: zipfile.ZipFile, member: str) -> Optional[ET.Element]:
    try:
        payload = read_capped(archive, member)
    except KeyError:
        return None
    try:
        return ET.fromstring(payload)
    except ET.ParseError as exc:
        raise ValueError(f"Invalid XML in {member}: {exc}") from exc


def enforce_uncompressed_limit(infos: list[zipfile.ZipInfo]) -> None:
    total = 0
    for info in infos:
        if info.file_size > MAX_UNCOMPRESSED_BYTES:
            raise ValueError(
                f"{info.filename} declares uncompressed size {info.file_size}, "
                f"limit is {MAX_UNCOMPRESSED_BYTES}"
            )
        total += info.file_size
        if total > MAX_UNCOMPRESSED_BYTES:
            raise ValueError(
                f"Uncompressed package size {total} exceeds {MAX_UNCOMPRESSED_BYTES}"
            )


def paragraph_text(paragraph: ET.Element) -> str:
    pieces: list[str] = []
    for node in iter_selected(paragraph):
        if node.tag == qn("w", "t"):
            pieces.append(node.text or "")
        elif node.tag == qn("m", "t"):
            pieces.append(node.text or "")
        elif node.tag == qn("w", "tab"):
            pieces.append("\t")
        elif node.tag in {qn("w", "br"), qn("w", "cr")}:
            pieces.append("\n")
    return "".join(pieces).strip()


def table_text(table: ET.Element) -> list[list[str]]:
    rows: list[list[str]] = []
    for row in selected_child_elements(table):
        if row.tag != qn("w", "tr"):
            continue
        values: list[str] = []
        for cell in selected_child_elements(row):
            if cell.tag != qn("w", "tc"):
                continue
            parts = []
            for node in iter_selected(cell):
                if node.tag == qn("w", "p"):
                    text = paragraph_text(node)
                    if text:
                        parts.append(text)
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


def chars_value(raw: Optional[str]) -> dict[str, Any] | None:
    """OOXML character indents are hundredths of a character (200 = 2 chars)."""
    if raw is None:
        return None
    try:
        value = int(raw)
    except ValueError:
        return {"raw": raw}
    return {"hundredths": value, "chars": round(value / 100, 2)}


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


def empty_font() -> dict[str, Any]:
    return {key: None for key in FONT_KEYS}


def empty_spacing() -> dict[str, Any]:
    return {
        "before": None,
        "after": None,
        "line": None,
        "line_rule": None,
        "line_multiple": None,
    }


def empty_indent() -> dict[str, Any]:
    return {
        "left": None,
        "right": None,
        "first_line": None,
        "hanging": None,
        "left_chars": None,
        "right_chars": None,
        "first_line_chars": None,
        "hanging_chars": None,
    }


def finalize_spacing(spacing: dict[str, Any]) -> dict[str, Any]:
    result = dict(spacing)
    rule = result.get("line_rule") or "auto"
    line = result.get("line")
    result["line_multiple"] = None
    if line is not None and str(rule) == "auto":
        try:
            result["line_multiple"] = round(int(line) / 240, 3)
        except (TypeError, ValueError):
            result["line_multiple"] = None
    return result


def spacing_from(element: Optional[ET.Element]) -> dict[str, Any]:
    spacing = empty_spacing()
    if element is None:
        return spacing
    spacing.update(
        {
            "before": twips_value(attr(element, "before")),
            "after": twips_value(attr(element, "after")),
            "line": attr(element, "line"),
            "line_rule": attr(element, "lineRule"),
        }
    )
    return finalize_spacing(spacing)


def indent_from(element: Optional[ET.Element]) -> dict[str, Any]:
    indent = empty_indent()
    if element is None:
        return indent
    indent.update(
        {
            "left": twips_value(attr(element, "left")),
            "right": twips_value(attr(element, "right")),
            "first_line": twips_value(attr(element, "firstLine")),
            "hanging": twips_value(attr(element, "hanging")),
            "left_chars": chars_value(attr(element, "leftChars")),
            "right_chars": chars_value(attr(element, "rightChars")),
            "first_line_chars": chars_value(attr(element, "firstLineChars")),
            "hanging_chars": chars_value(attr(element, "hangingChars")),
        }
    )
    return indent


def font_from(element: Optional[ET.Element], theme: dict[str, str]) -> dict[str, Any]:
    font = empty_font()
    if element is None:
        return font
    for key, explicit_name in EXPLICIT_FONT_ATTRS.items():
        explicit = attr(element, explicit_name)
        theme_name = attr(element, THEME_FONT_ATTRS[key])
        if explicit:
            font[key] = explicit
        elif theme_name and theme.get(theme_name):
            font[key] = theme[theme_name]
    return font


def merge_font(base: dict[str, Any], overlay: dict[str, Any]) -> dict[str, Any]:
    merged = dict(base)
    for key in FONT_KEYS:
        if overlay.get(key):
            merged[key] = overlay[key]
    return merged


def merge_mapping(base: dict[str, Any], overlay: dict[str, Any]) -> dict[str, Any]:
    merged = dict(base)
    for key, value in overlay.items():
        if value is not None:
            merged[key] = value
    return merged


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


def parse_theme(root: Optional[ET.Element]) -> dict[str, str]:
    if root is None:
        return {}

    def face(scheme: str, tag: str) -> Optional[str]:
        element = root.find(f".//a:fontScheme/a:{scheme}/a:{tag}", NS)
        if element is None:
            return None
        value = element.get("typeface")
        return value or None

    major_latin = face("majorFont", "latin")
    minor_latin = face("minorFont", "latin")
    major_ea = face("majorFont", "ea")
    minor_ea = face("minorFont", "ea")
    major_cs = face("majorFont", "cs")
    minor_cs = face("minorFont", "cs")
    mapping = {
        "majorAscii": major_latin,
        "majorHAnsi": major_latin,
        "majorEastAsia": major_ea,
        "majorBidi": major_cs,
        "minorAscii": minor_latin,
        "minorHAnsi": minor_latin,
        "minorEastAsia": minor_ea,
        "minorBidi": minor_cs,
    }
    return {key: value for key, value in mapping.items() if value}


def style_record(style: ET.Element, theme: dict[str, str]) -> dict[str, Any]:
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
        "font": font_from(fonts, theme),
        "size": half_points(attr(rpr.find("./w:sz", NS), "val") if rpr is not None else None),
        "size_cs": half_points(
            attr(rpr.find("./w:szCs", NS), "val") if rpr is not None else None
        ),
        "bold": on_off(rpr.find("./w:b", NS)) if rpr is not None else None,
        "italic": on_off(rpr.find("./w:i", NS)) if rpr is not None else None,
        "color": attr(color, "val"),
        "alignment": attr(justification, "val"),
        "spacing": spacing_from(spacing),
        "indent": indent_from(indent),
        "keep_next": on_off(ppr.find("./w:keepNext", NS)) if ppr is not None else None,
        "page_break_before": (
            on_off(ppr.find("./w:pageBreakBefore", NS)) if ppr is not None else None
        ),
        "inherited_properties": [],
    }


def level_from_outline(raw: Optional[str]) -> Optional[int]:
    """Map OOXML outlineLvl: 0-8 are headings 1-9, and 9 is body text."""
    if raw is None:
        return None
    try:
        value = int(raw)
    except ValueError:
        return None
    if 0 <= value <= 8:
        return value + 1
    return None


def heading_level_from_names(style_id: str, style_name: str) -> Optional[int]:
    for value in (style_id, style_name):
        compact = re.sub(r"\s+", "", value or "")
        match = HEADING_NAME_RE.search(compact)
        if match:
            return int(match.group(1))
    return None


def resolve_style_records(records: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    by_id = {record["style_id"]: record for record in records if record.get("style_id")}
    resolved: dict[str, dict[str, Any]] = {}
    resolving: set[str] = set()

    def resolve(style_id: Optional[str]) -> Optional[dict[str, Any]]:
        if not style_id or style_id not in by_id:
            return None
        if style_id in resolved:
            return resolved[style_id]
        if style_id in resolving:
            return None
        resolving.add(style_id)
        raw = by_id[style_id]
        parent = resolve(raw.get("based_on"))
        merged = {
            "style_id": raw["style_id"],
            "type": raw.get("type"),
            "name": raw.get("name"),
            "based_on": raw.get("based_on"),
            "next": raw.get("next"),
            "default": raw.get("default"),
            "outline_level": None,
            "font": empty_font(),
            "size": None,
            "size_cs": None,
            "bold": None,
            "italic": None,
            "color": None,
            "alignment": None,
            "spacing": empty_spacing(),
            "indent": empty_indent(),
            "keep_next": None,
            "page_break_before": None,
            "inherited_properties": [],
        }
        inherited: list[str] = []
        if parent is not None:
            for key in (
                "outline_level",
                "size",
                "size_cs",
                "bold",
                "italic",
                "color",
                "alignment",
                "keep_next",
                "page_break_before",
            ):
                if parent.get(key) is not None:
                    merged[key] = parent[key]
                    inherited.append(key)
            merged["font"] = dict(parent.get("font") or empty_font())
            merged["spacing"] = dict(parent.get("spacing") or empty_spacing())
            merged["indent"] = dict(parent.get("indent") or empty_indent())
            if any(merged["font"].values()):
                inherited.append("font")
            if any(value is not None for value in merged["spacing"].values()):
                inherited.append("spacing")
            if any(value is not None for value in merged["indent"].values()):
                inherited.append("indent")
        if raw.get("outline_level") is not None:
            merged["outline_level"] = raw["outline_level"]
            inherited = [item for item in inherited if item != "outline_level"]
        for key in ("size", "size_cs", "bold", "italic", "color", "alignment", "keep_next", "page_break_before"):
            if raw.get(key) is not None:
                merged[key] = raw[key]
                inherited = [item for item in inherited if item != key]
        if not font_is_empty(raw.get("font") or {}):
            merged["font"] = merge_font(merged["font"], raw["font"])
            inherited = [item for item in inherited if item != "font"]
        merged["spacing"] = finalize_spacing(merge_mapping(merged["spacing"], raw.get("spacing") or {}))
        if any((raw.get("spacing") or {}).get(key) is not None for key in ("before", "after", "line", "line_rule")):
            inherited = [item for item in inherited if item != "spacing"]
        merged["indent"] = merge_mapping(merged["indent"], raw.get("indent") or {})
        if any((raw.get("indent") or {}).get(key) is not None for key in empty_indent()):
            inherited = [item for item in inherited if item != "indent"]
        merged["inherited_properties"] = inherited
        merged["heading_level"] = heading_level_for_record(merged)
        resolving.remove(style_id)
        resolved[style_id] = merged
        return merged

    for style_id in by_id:
        resolve(style_id)
    return resolved


def font_is_empty(font: dict[str, Any]) -> bool:
    return not any(font.get(key) for key in FONT_KEYS)


def heading_level_for_record(record: dict[str, Any]) -> Optional[int]:
    if record.get("type") not in (None, "paragraph"):
        return None
    if record.get("outline_level") is not None:
        return level_from_outline(record.get("outline_level"))
    return heading_level_from_names(record.get("style_id") or "", record.get("name") or "")


def normalize_style_key(value: Optional[str]) -> str:
    return re.sub(r"\s+", "", value or "").lower()


def parse_styles(root: Optional[ET.Element], theme: dict[str, str]) -> dict[str, Any]:
    if root is None:
        return {
            "defaults": {},
            "selected": [],
            "outline_styles": [],
            "heading_style_map": {},
            "resolved": {},
        }

    defaults: dict[str, Any] = {}
    doc_defaults = root.find("./w:docDefaults", NS)
    if doc_defaults is not None:
        rpr = doc_defaults.find("./w:rPrDefault/w:rPr", NS)
        ppr = doc_defaults.find("./w:pPrDefault/w:pPr", NS)
        fonts = rpr.find("./w:rFonts", NS) if rpr is not None else None
        spacing = ppr.find("./w:spacing", NS) if ppr is not None else None
        indent = ppr.find("./w:ind", NS) if ppr is not None else None
        defaults = {
            "font": font_from(fonts, theme),
            "size": half_points(
                attr(rpr.find("./w:sz", NS), "val") if rpr is not None else None
            ),
            "spacing": spacing_from(spacing),
            "indent": indent_from(indent),
        }

    raw_records = [style_record(style, theme) for style in root.findall("./w:style", NS)]
    resolved = resolve_style_records(raw_records)
    heading_style_map = {
        style_id: record["heading_level"]
        for style_id, record in resolved.items()
        if record.get("heading_level") is not None
    }
    return {
        "defaults": defaults,
        "resolved": resolved,
        "heading_style_map": heading_style_map,
        "raw_count": len(raw_records),
    }


def select_styles(
    styles: dict[str, Any], used_style_ids: set[str]
) -> dict[str, Any]:
    resolved: dict[str, dict[str, Any]] = styles.get("resolved") or {}
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
        "正文",
    }
    selected: list[dict[str, Any]] = []
    outline_styles: list[dict[str, Any]] = []
    for record in resolved.values():
        style_id = normalize_style_key(record.get("style_id"))
        style_name = normalize_style_key(record.get("name"))
        is_heading = record.get("heading_level") is not None
        is_default = record.get("default") is True and record.get("type") in (None, "paragraph")
        is_used = record.get("style_id") in used_style_ids and record.get("type") in (
            None,
            "paragraph",
        )
        if style_id in wanted or style_name in wanted or is_heading or is_default or is_used:
            selected.append(record)
        outline = record.get("outline_level")
        if level_from_outline(outline) is not None:
            outline_styles.append(record)
    truncated = False
    if len(selected) > MAX_SELECTED_STYLES:
        selected = selected[:MAX_SELECTED_STYLES]
        truncated = True
    return {
        "defaults": styles.get("defaults") or {},
        "selected": selected,
        "outline_styles": outline_styles,
        "heading_style_map": styles.get("heading_style_map") or {},
        "truncated": truncated,
    }


def collect_fields(root: ET.Element) -> list[str]:
    values: list[str] = []
    for node in iter_selected(root):
        if node.tag == qn("w", "instrText"):
            value = " ".join((node.text or "").split())
        elif node.tag == qn("w", "fldSimple"):
            value = " ".join((attr(node, "instr") or "").split())
        else:
            continue
        if value:
            values.append(value)
    return values


def has_page_break(paragraph: ET.Element) -> bool:
    return any(
        node.tag == qn("w", "br") and attr(node, "type") == "page"
        for node in iter_selected(paragraph)
    )


def has_section_properties(paragraph: ET.Element) -> bool:
    return any(node.tag == qn("w", "sectPr") for node in iter_selected(paragraph))


def iter_runs(paragraph: ET.Element):
    def walk(element: ET.Element):
        for child in selected_child_elements(element):
            if child.tag == qn("w", "txbxContent"):
                continue
            if child.tag == qn("w", "r"):
                yield child
            else:
                yield from walk(child)

    yield from walk(paragraph)


def paragraph_direct(paragraph: ET.Element, theme: dict[str, str]) -> dict[str, Any]:
    ppr = first_selected(paragraph, "pPr")
    direct: dict[str, Any] = {}
    alignment = attr(first_selected(ppr, "jc"), "val")
    if alignment:
        direct["alignment"] = alignment
    spacing = spacing_from(first_selected(ppr, "spacing"))
    if any(spacing.get(key) is not None for key in ("before", "after", "line", "line_rule")):
        direct["spacing"] = spacing
    indent = indent_from(first_selected(ppr, "ind"))
    if any(value is not None for value in indent.values()):
        direct["indent"] = indent
    fonts = []
    sizes = []
    bolds = []
    italics = []
    for node in iter_runs(paragraph):
        rpr = first_selected(node, "rPr")
        if rpr is None:
            continue
        font = font_from(first_selected(rpr, "rFonts"), theme)
        if any(font.values()):
            fonts.append(font)
        size = half_points(attr(first_selected(rpr, "sz"), "val"))
        if size is not None:
            sizes.append(size)
        bold = on_off(first_selected(rpr, "b"))
        if bold is not None:
            bolds.append(bold)
        italic = on_off(first_selected(rpr, "i"))
        if italic is not None:
            italics.append(italic)
    if fonts:
        direct["font"] = fonts[0]
        if any(font != fonts[0] for font in fonts[1:]):
            direct["mixed_runs"] = True
    if sizes:
        direct["size"] = sizes[0]
    if bolds:
        direct["bold"] = bolds[0]
    if italics:
        direct["italic"] = italics[0]
    return direct


def apply_direct(base: dict[str, Any], direct: dict[str, Any]) -> dict[str, Any]:
    effective = {
        "font": dict(base.get("font") or empty_font()),
        "size": base.get("size"),
        "bold": base.get("bold"),
        "italic": base.get("italic"),
        "alignment": base.get("alignment"),
        "spacing": dict(base.get("spacing") or empty_spacing()),
        "indent": dict(base.get("indent") or empty_indent()),
    }
    if direct.get("font"):
        effective["font"] = merge_font(effective["font"], direct["font"])
    for key in ("size", "bold", "italic", "alignment"):
        if direct.get(key) is not None:
            effective[key] = direct[key]
    if direct.get("spacing"):
        effective["spacing"] = finalize_spacing(
            merge_mapping(effective["spacing"], direct["spacing"])
        )
    if direct.get("indent"):
        effective["indent"] = merge_mapping(effective["indent"], direct["indent"])
    return effective


def paragraph_record(
    paragraph: ET.Element,
    index: int,
    container: str,
    style_heading_map: dict[str, int],
    resolved_styles: dict[str, dict[str, Any]],
    numbering_defs: dict[str, dict[str, dict[str, Any]]],
    theme: dict[str, str],
) -> dict[str, Any]:
    ppr = first_selected(paragraph, "pPr")
    pstyle = first_selected(ppr, "pStyle")
    outline = first_selected(ppr, "outlineLvl")
    numpr = first_selected(ppr, "numPr")
    num_id_element = first_selected(numpr, "numId")
    ilvl = first_selected(numpr, "ilvl")
    text = paragraph_text(paragraph)
    style_id = attr(pstyle, "val")
    direct_outline = attr(outline, "val")

    heading_level: Optional[int] = None
    if direct_outline is not None:
        heading_level = level_from_outline(direct_outline)
    elif style_id:
        heading_level = style_heading_map.get(style_id)

    heading_pattern = None
    patterns = [
        ("chapter_cn", r"^第[一二三四五六七八九十百0-9]+章(?:\s|[：:、.]|$)", None),
        ("decimal", r"^\d+\.\d+(?:\.\d+){0,3}(?:\s|[、.]|$)", None),
        ("decimal", r"^\d{1,2}(?:\s|[、.])\S+$", 40),
        ("chinese_list", r"^[一二三四五六七八九十]+、", None),
    ]
    for name, pattern, max_length in patterns:
        if max_length is not None and len(text) > max_length:
            continue
        if name == "decimal" and re.search(r"[。；;!?]", text):
            continue
        if re.search(pattern, text):
            heading_pattern = name
            break

    num_id = attr(num_id_element, "val")
    level = attr(ilvl, "val")
    numbering: dict[str, Any] = {
        "num_id": num_id,
        "level": level,
        "enabled": bool(num_id) and num_id != "0",
    }
    if numbering["enabled"] and num_id in numbering_defs:
        definition = numbering_defs[num_id].get(level or "0")
        if definition:
            numbering.update(definition)

    base = resolved_styles.get(style_id or "") or {}
    direct = paragraph_direct(paragraph, theme)
    return {
        "index": index,
        "container": container,
        "text": text,
        "style_id": style_id,
        "outline_level": direct_outline if direct_outline is not None else base.get("outline_level"),
        "heading_level": heading_level,
        "heading_pattern": heading_pattern,
        "numbering": numbering,
        "page_break": has_page_break(paragraph),
        "section_break": has_section_properties(paragraph),
        "direct_formatting": direct,
        "effective": apply_direct(base, direct),
    }


def section_record(section: ET.Element, index: int) -> dict[str, Any]:
    size = first_selected(section, "pgSz")
    margins = first_selected(section, "pgMar")
    number = first_selected(section, "pgNumType")
    section_type = first_selected(section, "type")
    columns = first_selected(section, "cols")
    title_page = first_selected(section, "titlePg")

    refs = []
    for ref in selected_child_elements(section):
        if ref.tag not in {qn("w", "headerReference"), qn("w", "footerReference")}:
            continue
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
    archive: zipfile.ZipFile,
    relationships: dict[str, str],
    warnings: list[dict[str, Any]],
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
        try:
            root = read_xml(archive, normalized)
        except ValueError as exc:
            warnings.append(
                {
                    "code": "unreadable_part",
                    "part": normalized,
                    "relationship_id": rel_id,
                    "message": str(exc),
                }
            )
            continue
        if root is None:
            continue
        text = " ".join(
            part
            for part in (
                paragraph_text(child)
                for child in selected_child_elements(root)
                if child.tag == qn("w", "p")
            )
            if part
        )
        records.append(
            {
                "relationship_id": rel_id,
                "part": normalized,
                "kind": "header" if basename.startswith("header") else "footer",
                "text": text,
                "fields": " ".join(collect_fields(root)),
            }
        )
    return records


def collect_content(
    body: ET.Element, warnings: list[dict[str, Any]]
) -> tuple[list[tuple[ET.Element, str]], list[tuple[int, ET.Element]]]:
    paragraphs: list[tuple[ET.Element, str]] = []
    tables: list[tuple[int, ET.Element]] = []

    def walk(element: ET.Element, container: str) -> None:
        for child in selected_child_elements(element):
            if child.tag == qn("w", "p"):
                paragraphs.append((child, container))
            elif child.tag == qn("w", "tbl"):
                tables.append((len(paragraphs), child))
                walk(child, "table")
            elif child.tag == qn("w", "sdt"):
                content = first_selected(child, "sdtContent")
                if content is not None:
                    next_container = "sdt" if container == "body" else container
                    walk(content, next_container)
            elif child.tag == qn("w", "altChunk"):
                warnings.append(
                    {
                        "code": "alt_chunk_not_expanded",
                        "relationship_id": attr(child, "id", "r"),
                        "message": (
                            "altChunk content is not expanded. Profile text does not "
                            "include the alternate chunk."
                        ),
                    }
                )
            elif child.tag in {
                qn("w", "tr"),
                qn("w", "tc"),
                qn("w", "customXml"),
                qn("w", "smartTag"),
                qn("w", "hyperlink"),
                qn("w", "ins"),
                qn("w", "sdtContent"),
            }:
                walk(child, container)

    walk(body, "body")
    return paragraphs, tables


def cover_candidates(
    paragraphs: list[dict[str, Any]],
    tables: list[tuple[int, ET.Element]],
) -> dict[str, Any]:
    boundaries = []
    for paragraph in paragraphs:
        if paragraph.get("heading_level") in range(1, 10) and paragraph.get("text"):
            boundaries.append(paragraph["index"])
            break
    for paragraph in paragraphs:
        if paragraph.get("page_break") or paragraph.get("section_break"):
            boundaries.append(paragraph["index"])
            break
    boundary = min(boundaries) if boundaries else None
    if boundary is None:
        paragraph_limit = min(len(paragraphs), MAX_COVER_PARAGRAPHS)
        selected = [
            paragraph
            for paragraph in paragraphs
            if paragraph["index"] < paragraph_limit and paragraph.get("text")
        ][:MAX_COVER_PARAGRAPHS]
        selected_tables = [table for anchor, table in tables if anchor < paragraph_limit]
        cutoff = paragraph_limit - 1 if paragraphs else None
    else:
        selected = [
            paragraph
            for paragraph in paragraphs
            if paragraph["index"] < boundary and paragraph.get("text")
        ][:MAX_COVER_PARAGRAPHS]
        selected_tables = [table for anchor, table in tables if anchor <= boundary]
        cutoff = boundary

    return {
        "cutoff_paragraph_index": cutoff,
        "paragraphs": selected,
        "tables": [table_text(table) for table in selected_tables[:MAX_COVER_TABLES]],
    }


def settings_record(root: Optional[ET.Element]) -> dict[str, Any]:
    """Return settings that materially affect field refresh and pagination."""
    if root is None:
        return {
            "update_fields_on_open": False,
            "even_and_odd_headers": False,
            "mirror_margins": False,
            "doc_grid": None,
        }
    grid = root.find("./w:docGrid", NS)
    doc_grid = None
    if grid is not None:
        pitch = attr(grid, "linePitch")
        pitch_pt = None
        if pitch is not None:
            try:
                pitch_pt = round(int(pitch) / 20, 2)
            except ValueError:
                pitch_pt = None
        doc_grid = {
            "type": attr(grid, "type"),
            "line_pitch": twips_value(pitch),
            "line_pitch_pt": pitch_pt,
            "char_space": attr(grid, "charSpace"),
        }
    return {
        "update_fields_on_open": on_off(root.find("./w:updateFields", NS)) is True,
        "even_and_odd_headers": on_off(root.find("./w:evenAndOddHeaders", NS)) is True,
        "mirror_margins": on_off(root.find("./w:mirrorMargins", NS)) is True,
        "doc_grid": doc_grid,
    }


def pagination_warnings(
    settings: dict[str, Any], sections: list[dict[str, Any]]
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


def parse_numbering(
    root: Optional[ET.Element],
) -> tuple[dict[str, dict[str, dict[str, Any]]], list[dict[str, Any]]]:
    if root is None:
        return {}, []
    abstracts: dict[str, dict[str, dict[str, Any]]] = {}
    for abstract in root.findall("./w:abstractNum", NS):
        abstract_id = attr(abstract, "abstractNumId")
        if not abstract_id:
            continue
        levels: dict[str, dict[str, Any]] = {}
        for level in abstract.findall("./w:lvl", NS):
            ilvl = attr(level, "ilvl") or "0"
            levels[ilvl] = {
                "format": attr(level.find("./w:numFmt", NS), "val"),
                "text": attr(level.find("./w:lvlText", NS), "val"),
                "start": attr(level.find("./w:start", NS), "val"),
                "style_id": attr(level.find("./w:pStyle", NS), "val"),
            }
        abstracts[abstract_id] = levels
    nums: dict[str, dict[str, dict[str, Any]]] = {}
    definitions: list[dict[str, Any]] = []
    for num in root.findall("./w:num", NS)[:50]:
        num_id = attr(num, "numId")
        abstract_id = attr(num.find("./w:abstractNumId", NS), "val")
        if not num_id:
            continue
        levels = dict(abstracts.get(abstract_id or "", {}))
        nums[num_id] = levels
        definitions.append(
            {
                "num_id": num_id,
                "abstract_num_id": abstract_id,
                "levels": [
                    {"level": level, **payload}
                    for level, payload in sorted(
                        levels.items(),
                        key=lambda item: int(item[0]) if item[0].isdigit() else 0,
                    )
                ],
            }
        )
    return nums, definitions


def count_notes(root: Optional[ET.Element], local: str) -> int:
    if root is None:
        return 0
    count = 0
    for note in root.findall(f"./w:{local}", NS):
        note_type = attr(note, "type")
        if note_type in NOTE_SKIP_TYPES:
            continue
        count += 1
    return count


def vba_project_present(archive: zipfile.ZipFile) -> bool:
    names = set(archive.namelist())
    if "word/vbaProject.bin" in names:
        return True
    try:
        content_types = archive.read("[Content_Types].xml")
    except KeyError:
        return False
    return b"vbaProject" in content_types


def profile(path: Path) -> dict[str, Any]:
    if path.suffix.lower() not in {".docx", ".dotx", ".docm", ".dotm"}:
        raise ValueError("Input must be a DOCX, DOTX, DOCM, or DOTM file")
    if not path.exists():
        raise FileNotFoundError(path)
    size_bytes = path.stat().st_size
    if size_bytes > MAX_PACKAGE_BYTES:
        raise ValueError(
            f"Package size {size_bytes} exceeds {MAX_PACKAGE_BYTES} bytes"
        )

    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    with zipfile.ZipFile(path) as archive:
        enforce_uncompressed_limit(archive.infolist())
        document = read_xml(archive, "word/document.xml")
        if document is None:
            raise ValueError("word/document.xml is missing")
        body = document.find("./w:body", NS)
        if body is None:
            raise ValueError("Document body is missing")

        theme = parse_theme(read_xml(archive, "word/theme/theme1.xml"))
        styles_raw = parse_styles(read_xml(archive, "word/styles.xml"), theme)
        resolved_styles = styles_raw.get("resolved") or {}
        settings = settings_record(read_xml(archive, "word/settings.xml"))
        numbering_defs, numbering_definitions = parse_numbering(
            read_xml(archive, "word/numbering.xml")
        )
        rels = parse_relationships(read_xml(archive, "word/_rels/document.xml.rels"))
        warnings = pagination_warnings(settings, [])
        part_warnings: list[dict[str, Any]] = []

        paragraph_elements, table_anchors = collect_content(body, part_warnings)
        if len(paragraph_elements) > MAX_PARAGRAPHS:
            raise ValueError(
                f"Paragraph count {len(paragraph_elements)} exceeds {MAX_PARAGRAPHS}"
            )
        paragraphs = [
            paragraph_record(
                element,
                index,
                container,
                styles_raw.get("heading_style_map") or {},
                resolved_styles,
                numbering_defs,
                theme,
            )
            for index, (element, container) in enumerate(paragraph_elements)
        ]
        used_style_ids = {paragraph["style_id"] for paragraph in paragraphs if paragraph.get("style_id")}
        styles = select_styles(styles_raw, used_style_ids)
        if styles.pop("truncated", False):
            part_warnings.append(
                {
                    "code": "style_list_truncated",
                    "message": f"Selected style list was truncated to {MAX_SELECTED_STYLES}.",
                }
            )
        sections = [
            section_record(section, index)
            for index, section in enumerate(
                node for node in iter_selected(document) if node.tag == qn("w", "sectPr")
            )
        ]
        warnings = pagination_warnings(settings, sections)
        warnings.extend(part_warnings)
        fields = collect_fields(document)
        names = set(archive.namelist())
        macro_extension = path.suffix.lower() in {".docm", ".dotm"}
        has_vba = vba_project_present(archive)

        return {
            "schema_version": 1,
            "file": {
                "path": str(path),
                "name": path.name,
                "size_bytes": size_bytes,
                "sha256": digest,
                "macro_enabled": macro_extension or has_vba,
                "vba_project_present": has_vba,
            },
            "counts": {
                "paragraphs": len(paragraphs),
                "nonempty_paragraphs": sum(bool(p["text"]) for p in paragraphs),
                "tables": len(table_anchors),
                "images": len(
                    [
                        name
                        for name in names
                        if name.startswith("word/media/") and not name.endswith("/")
                    ]
                ),
                "sections": len(sections),
                "comments": count_notes(read_xml(archive, "word/comments.xml"), "comment"),
                "footnotes": count_notes(read_xml(archive, "word/footnotes.xml"), "footnote"),
                "endnotes": count_notes(read_xml(archive, "word/endnotes.xml"), "endnote"),
            },
            "styles": styles,
            "numbering_definitions": numbering_definitions,
            "settings": settings,
            "sections": sections,
            "warnings": warnings,
            "headers_footers": header_footer_profiles(archive, rels, warnings),
            "fields": fields,
            "has_toc_field": any(
                re.search(r"(?:^|\s)TOC(?:\s|$)", field, re.IGNORECASE)
                for field in fields
            ),
            "paragraphs": paragraphs,
            "cover_candidates": cover_candidates(paragraphs, table_anchors),
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
