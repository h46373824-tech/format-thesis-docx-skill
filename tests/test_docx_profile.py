"""Synthetic OOXML regression tests for the thesis profiler.

Fixtures are generated in a temporary directory. They are not real Word
pagination samples and do not prove desktop Word layout.
"""

from __future__ import annotations

import importlib.util
import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "format-thesis-docx" / "scripts" / "docx_profile.py"
SPEC = importlib.util.spec_from_file_location("docx_profile", SCRIPT)
assert SPEC and SPEC.loader
profile_mod = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(profile_mod)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def package(path: Path, document_xml: str, extra: dict | None = None, rels_xml: str | None = None) -> Path:
    extra = extra or {}
    content_types = f"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="{CT}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""
    root_rels = f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="{PR}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("word/document.xml", document_xml)
        if rels_xml is not None:
            archive.writestr("word/_rels/document.xml.rels", rels_xml)
        for name, payload in extra.items():
            archive.writestr(name, payload)
    return path


def document(body: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:m="{M}">'
        f"<w:body>{body}</w:body></w:document>"
    )


def paragraph(text: str, style: str | None = None, ppr: str = "", rpr: str = "") -> str:
    style_xml = f'<w:pStyle w:val="{style}"/>' if style else ""
    properties = f"<w:pPr>{style_xml}{ppr}</w:pPr>" if style or ppr else ""
    run_properties = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
    return (
        f"<w:p>{properties}<w:r>{run_properties}"
        f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    )


class DocxProfileTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.directory = Path(self.temp_dir.name)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_chinese_layout_inheritance_controls_and_fields(self) -> None:
        styles = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="{W}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi"/>
      <w:sz w:val="24"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>
      <w:ind w:firstLineChars="200"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="a" w:default="1">
    <w:name w:val="正文"/>
    <w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman"/></w:rPr>
    <w:pPr><w:spacing w:line="360" w:lineRule="auto"/><w:ind w:firstLineChars="200"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="1">
    <w:name w:val="标题 1"/>
    <w:basedOn w:val="a"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:sz w:val="32"/><w:b/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ChapterTitle">
    <w:name w:val="章标题"/>
    <w:basedOn w:val="1"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading10">
    <w:name w:val="Heading 10"/>
    <w:pPr><w:outlineLvl w:val="8"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="BodyText9">
    <w:name w:val="正文无大纲"/>
    <w:pPr><w:outlineLvl w:val="9"/></w:pPr>
  </w:style>
  <w:style w:type="character" w:styleId="Heading1Char">
    <w:name w:val="Heading 1 Char"/>
  </w:style>
</w:styles>
"""
        theme = """<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
  <a:themeElements><a:fontScheme name="Office">
    <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface="黑体"/></a:majorFont>
    <a:minorFont><a:latin typeface="Times New Roman"/><a:ea typeface="宋体"/></a:minorFont>
  </a:fontScheme></a:themeElements>
</a:theme>
"""
        settings = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:settings xmlns:w="{W}">
  <w:mirrorMargins/>
  <w:evenAndOddHeaders/>
  <w:updateFields w:val="true"/>
  <w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0"/>
</w:settings>
"""
        numbering = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="{W}">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="chineseCounting"/>
      <w:lvlText w:val="第%1章"/>
      <w:pStyle w:val="1"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
"""
        footnotes = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:footnotes xmlns:w="{W}">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:r><w:t>合成脚注</w:t></w:r></w:p></w:footnote>
</w:footnotes>
"""
        header = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>页眉学校名</w:t></w:r></w:p>
<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>I</w:t></w:r></w:fldSimple></w:p></w:hdr>
"""
        footer = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:ftr xmlns:w="{W}"><w:p><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r></w:p></w:ftr>
"""
        rels = f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="{PR}">
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>
"""
        body = (
            '<w:sdt><w:sdtContent><w:p><w:r><w:t>内容控件中的题目</w:t></w:r></w:p>'
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>控件表</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            "</w:sdtContent></w:sdt>"
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>学号</w:t></w:r></w:p></w:tc>"
            "<w:tc><w:p><w:r><w:t>20260001</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            + paragraph("封面题目", ppr='<w:jc w:val="center"/>')
            + paragraph(
                "作者甲",
                rpr='<w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体"/><w:sz w:val="30"/>',
            )
            + '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            + '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1701"/>'
            + '<w:pgNumType w:fmt="upperRoman" w:start="1"/></w:sectPr></w:pPr></w:p>'
            + '<w:p><w:fldSimple w:instr=" TOC \\o &quot;1-3&quot; \\h "><w:r><w:t>目录占位</w:t></w:r></w:fldSimple></w:p>'
            + f'<w:p><w:r><m:oMath xmlns:m="{M}"><m:r><m:t>E=mc^2</m:t></m:r></m:oMath></w:r></w:p>'
            + paragraph("第一章 绪论", "1")
            + paragraph("继承章标题", "ChapterTitle")
            + paragraph("自定义十级", "Heading10")
            + paragraph("大纲九正文", "BodyText9")
            + paragraph(
                "直接格式正文",
                ppr='<w:spacing w:line="480" w:lineRule="exact"/><w:ind w:firstLineChars="200"/><w:jc w:val="both"/>',
            )
            + paragraph("编号段落", ppr='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>')
            + paragraph("1. 这是正文编号句，不是标题。")
            + paragraph("1. Smith J. A synthetic reference entry.")
            + paragraph("1.1 研究背景")
            + '<w:altChunk r:id="alt1"/>'
            + '<w:sectPr><w:type w:val="oddPage"/><w:pgSz w:w="11906" w:h="16838"/>'
            + '<w:pgNumType w:fmt="decimal" w:start="1"/></w:sectPr>'
        )
        path = package(
            self.directory / "case_layout.docx",
            document(body),
            extra={
                "word/styles.xml": styles,
                "word/settings.xml": settings,
                "word/numbering.xml": numbering,
                "word/theme/theme1.xml": theme,
                "word/header1.xml": header,
                "word/footer1.xml": footer,
                "word/footnotes.xml": footnotes,
                "word/vbaProject.bin": b"synthetic-vba",
            },
            rels_xml=rels,
        )
        result = profile_mod.profile(path)
        blob = json.dumps(result, ensure_ascii=False)
        self.assertIn("宋体", blob)
        self.assertIn("黑体", blob)
        self.assertEqual(result["file"]["path"], str(path))
        self.assertNotIn(str(path.resolve()) + "/resolved-only", result["file"]["path"])
        self.assertTrue(result["file"]["vba_project_present"])
        self.assertTrue(result["file"]["macro_enabled"])
        self.assertEqual(result["counts"]["footnotes"], 1)
        self.assertEqual(result["counts"]["tables"], 2)
        self.assertTrue(result["has_toc_field"])
        self.assertTrue(any(field.strip().startswith("TOC") for field in result["fields"]))
        self.assertEqual(result["styles"]["defaults"]["font"]["east_asia"], "宋体")
        self.assertEqual(result["styles"]["defaults"]["spacing"]["line_multiple"], 1.5)
        self.assertEqual(result["styles"]["defaults"]["indent"]["first_line_chars"]["chars"], 2)
        self.assertTrue(result["settings"]["mirror_margins"])
        self.assertEqual(result["settings"]["doc_grid"]["line_pitch_pt"], 15.6)
        selected = {item["style_id"]: item for item in result["styles"]["selected"]}
        self.assertEqual(selected["a"]["font"]["east_asia"], "宋体")
        self.assertEqual(selected["1"]["font"]["east_asia"], "宋体")
        self.assertEqual(selected["1"]["heading_level"], 1)
        self.assertNotIn("Heading1Char", result["styles"]["heading_style_map"])
        self.assertEqual(result["styles"]["heading_style_map"]["Heading10"], 9)
        self.assertNotIn("BodyText9", result["styles"]["heading_style_map"])
        by_text = {item["text"]: item for item in result["paragraphs"]}
        self.assertEqual(by_text["内容控件中的题目"]["container"], "sdt")
        self.assertEqual(by_text["控件表"]["container"], "table")
        self.assertEqual(by_text["学号"]["container"], "table")
        self.assertEqual(by_text["作者甲"]["effective"]["font"]["east_asia"], "黑体")
        self.assertEqual(by_text["作者甲"]["effective"]["size"]["pt"], 15)
        self.assertEqual(by_text["封面题目"]["effective"]["alignment"], "center")
        self.assertIn("E=mc^2", by_text["E=mc^2"]["text"])
        self.assertEqual(by_text["第一章 绪论"]["heading_level"], 1)
        self.assertEqual(by_text["第一章 绪论"]["effective"]["font"]["east_asia"], "宋体")
        self.assertEqual(by_text["继承章标题"]["heading_level"], 1)
        self.assertEqual(by_text["自定义十级"]["heading_level"], 9)
        self.assertIsNone(by_text["大纲九正文"]["heading_level"])
        self.assertIsNone(by_text["1. 这是正文编号句，不是标题。"]["heading_pattern"])
        self.assertIsNone(by_text["1. Smith J. A synthetic reference entry."]["heading_pattern"])
        self.assertEqual(by_text["1.1 研究背景"]["heading_pattern"], "decimal")
        self.assertFalse(by_text["编号段落"]["numbering"]["enabled"])
        self.assertEqual(by_text["直接格式正文"]["effective"]["indent"]["first_line_chars"]["chars"], 2)
        self.assertIsNone(by_text["直接格式正文"]["effective"]["spacing"]["line_multiple"])
        self.assertEqual(by_text["直接格式正文"]["effective"]["spacing"]["line_rule"], "exact")
        cover_text = [item["text"] for item in result["cover_candidates"]["paragraphs"]]
        self.assertIn("内容控件中的题目", cover_text)
        self.assertIn("学号", cover_text)
        self.assertIn("封面题目", cover_text)
        self.assertNotIn("第一章 绪论", cover_text)
        self.assertNotIn("目录占位", cover_text)
        flat_tables = json.dumps(result["cover_candidates"]["tables"], ensure_ascii=False)
        self.assertIn("20260001", flat_tables)
        self.assertIn("第%1章", json.dumps(result["numbering_definitions"], ensure_ascii=False))
        self.assertEqual(result["numbering_definitions"][0]["levels"][0]["format"], "chineseCounting")
        header = next(item for item in result["headers_footers"] if item["kind"] == "header")
        footer = next(item for item in result["headers_footers"] if item["kind"] == "footer")
        self.assertIn("PAGE", header["fields"])
        self.assertIn("PAGE", footer["fields"])
        self.assertIn("页眉学校名", header["text"])
        self.assertTrue(any(item["code"] == "alt_chunk_not_expanded" for item in result["warnings"]))
        self.assertTrue(any(item["code"] == "possible_virtual_even_page" for item in result["warnings"]))
        self.assertEqual(result["counts"]["sections"], 2)

    def test_alternate_content_is_not_double_counted(self) -> None:
        body = (
            f'<w:p><w:r><mc:AlternateContent xmlns:mc="{MC}">'
            f'<mc:Choice Requires="w14"><w:t>选择文本</w:t></mc:Choice>'
            f'<mc:Fallback><w:t>回退文本</w:t></mc:Fallback>'
            "</mc:AlternateContent></w:r></w:p>"
            f'<w:sectPr><mc:AlternateContent xmlns:mc="{MC}">'
            '<mc:Choice Requires="w14"><w:pgSz w:w="11906" w:h="16838"/></mc:Choice>'
            '<mc:Fallback><w:pgSz w:w="12240" w:h="15840"/></mc:Fallback>'
            "</mc:AlternateContent></w:sectPr>"
        )
        path = package(self.directory / "alternate.docx", document(body))
        result = profile_mod.profile(path)
        self.assertEqual([item["text"] for item in result["paragraphs"]], ["回退文本"])
        self.assertEqual(result["counts"]["sections"], 1)
        self.assertEqual(result["sections"][0]["page_size"]["width"]["twips"], 12240)

    def test_unreadable_header_does_not_fail_profile(self) -> None:
        rels = f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="{PR}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="media/header.png"/>
</Relationships>
"""
        path = package(
            self.directory / "bad-header.docx",
            document(paragraph("仅正文")),
            extra={"word/media/header.png": b"\x89PNG\r\nnot-xml"},
            rels_xml=rels,
        )
        result = profile_mod.profile(path)
        self.assertEqual(result["paragraphs"][0]["text"], "仅正文")
        self.assertTrue(any(item["code"] == "unreadable_part" for item in result["warnings"]))
        completed = subprocess.run(
            [sys.executable, str(SCRIPT), str(path)],
            check=False,
            capture_output=True,
            text=True,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)

    def test_entities_and_resource_limits_are_rejected(self) -> None:
        entity = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE w:document [<!ENTITY a "AAAA"><!ENTITY b "&a;&a;">]>
<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>&b;</w:t></w:r></w:p></w:body></w:document>
"""
        entity_path = package(self.directory / "entity.docx", entity)
        with self.assertRaises(ValueError):
            profile_mod.profile(entity_path)
        completed = subprocess.run(
            [sys.executable, str(SCRIPT), str(entity_path)],
            check=False,
            capture_output=True,
            text=True,
        )
        self.assertEqual(completed.returncode, 1)
        self.assertIn("entity", completed.stderr.lower())

        many = package(
            self.directory / "many.docx",
            document("".join(paragraph(f"段{index}") for index in range(4))),
        )
        original = profile_mod.MAX_PARAGRAPHS
        profile_mod.MAX_PARAGRAPHS = 3
        try:
            with self.assertRaises(ValueError):
                profile_mod.profile(many)
        finally:
            profile_mod.MAX_PARAGRAPHS = original

        original_xml = profile_mod.MAX_XML_BYTES
        profile_mod.MAX_XML_BYTES = 32
        try:
            with self.assertRaises(ValueError):
                profile_mod.profile(many)
        finally:
            profile_mod.MAX_XML_BYTES = original_xml

    def test_macro_extension_and_relative_path(self) -> None:
        dotm = package(self.directory / "plain.dotm", document(paragraph("模板")))
        dotm_result = profile_mod.profile(dotm)
        self.assertTrue(dotm_result["file"]["macro_enabled"])
        self.assertFalse(dotm_result["file"]["vba_project_present"])

        docx = package(self.directory / "plain.docx", document(paragraph("正文")))
        docx_result = profile_mod.profile(docx)
        self.assertFalse(docx_result["file"]["macro_enabled"])
        self.assertFalse(docx_result["file"]["vba_project_present"])

        previous = Path.cwd()
        os.chdir(self.directory)
        try:
            package(Path("relative-case.docx"), document(paragraph("相对路径")))
            result = profile_mod.profile(Path("relative-case.docx"))
            self.assertEqual(result["file"]["path"], "relative-case.docx")
            self.assertFalse(Path(result["file"]["path"]).is_absolute())
        finally:
            os.chdir(previous)

    def test_english_heading_style_still_maps(self) -> None:
        styles = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="{W}">
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
  </w:style>
</w:styles>
"""
        path = package(
            self.directory / "english.docx",
            document(paragraph("Introduction", "Heading1")),
            extra={"word/styles.xml": styles},
        )
        result = profile_mod.profile(path)
        self.assertEqual(result["styles"]["heading_style_map"]["Heading1"], 1)
        selected = {item["style_id"]: item for item in result["styles"]["selected"]}
        self.assertEqual(selected["Heading1"]["font"]["ascii"], "Calibri")
        self.assertEqual(result["paragraphs"][0]["heading_level"], 1)


if __name__ == "__main__":
    unittest.main()
