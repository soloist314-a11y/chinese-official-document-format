#!/usr/bin/env python3
"""Meaningful XML and input-safety regression checks; no rendering claims."""

import copy
import tempfile
import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET

import format_docx as formatter


class DailyFormatTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.output = Path(self.temp.name) / "result.docx"
        self.spec = {
            "mode": "daily",
            "blocks": [
                {"role": "title", "text": "格式检验"},
                {"role": "recipient", "text": "有关单位："},
                {"role": "h1", "text": "一、手工编号"},
                {"role": "body", "text": "保留原文 & <标记>，含换行\n第二行\t制表符。"},
                {"role": "table", "rows": [["项目", "内容"], ["示例", "不设固定行高"]]},
                {"role": "body", "text": "表后正文。"},
                {"role": "attachment_list", "items": [{"text": "附件：资料清单", "continuation_indent_chars": 5}]},
                {"role": "signature", "text": "发文主体", "align_with_date": "2026年10月2日"},
                {"role": "date", "text": "2026年10月2日"},
                {"role": "appendix_marker", "text": "附件"},
                {"role": "body", "text": "附件内容。"},
            ],
        }

    def tearDown(self):
        self.temp.cleanup()

    def document(self):
        with ZipFile(self.output) as z:
            return ET.fromstring(z.read("word/document.xml"))

    def rewrite_part(self, name, update):
        with ZipFile(self.output) as archive:
            parts = {part: archive.read(part) for part in archive.namelist()}
        tree = ET.fromstring(parts[name])
        update(tree)
        parts[name] = formatter.xml_bytes(tree)
        with ZipFile(self.output, "w", ZIP_DEFLATED) as archive:
            for part, data in parts.items():
                archive.writestr(part, data)

    def test_output_preserves_text_and_structure(self):
        report = formatter.create(self.spec, self.output)
        self.assertTrue(report["passed"], report["errors"])
        doc = self.document()
        self.assertEqual(len(doc.findall(".//w:numPr", formatter.NS)), 0)
        self.assertEqual(len(doc.findall(".//w:trHeight", formatter.NS)), 0)
        self.assertTrue(doc.findall(".//w:tblHeader", formatter.NS))
        self.assertTrue(doc.findall(".//w:pageBreakBefore", formatter.NS))
        body_para = doc.findall("w:body/w:p", formatter.NS)[3]
        result = []
        for node in body_para.iter():
            if node.tag == formatter.wtag("t"):
                result.append(node.text or "")
            elif node.tag == formatter.wtag("br"):
                result.append("\n")
            elif node.tag == formatter.wtag("tab"):
                result.append("\t")
        self.assertEqual("".join(result), self.spec["blocks"][3]["text"])
        margin = doc.find("w:body/w:sectPr/w:pgMar", formatter.NS)
        self.assertEqual(formatter.att(margin, "top"), "2098")
        grid = doc.find("w:body/w:sectPr/w:docGrid", formatter.NS)
        self.assertEqual(formatter.att(grid, "linePitch"), "585")
        for metadata in ("docProps/core.xml", "docProps/app.xml"):
            with ZipFile(self.output) as z:
                self.assertFalse(list(ET.fromstring(z.read(metadata))))

    def test_tampered_format_and_automatic_numbering_fail(self):
        formatter.create(self.spec, self.output)
        def mutate(doc):
            p = doc.findall("w:body/w:p", formatter.NS)[3]
            formatter.sub(p.find("w:pPr", formatter.NS), "numPr")
            margin = doc.find("w:body/w:sectPr/w:pgMar", formatter.NS)
            margin.set(formatter.wtag("top"), "1000")
            font = p.find("w:r/w:rPr/w:rFonts", formatter.NS)
            font.set(formatter.wtag("eastAsia"), "宋体")
        self.rewrite_part("word/document.xml", mutate)
        report = formatter.check(self.output)
        self.assertFalse(report["passed"])
        self.assertTrue(any("Automatic numbering" in x for x in report["errors"]))
        self.assertTrue(any("top margin" in x for x in report["errors"]))
        self.assertTrue(any("font differs" in x for x in report["errors"]))

    def test_metadata_leak_fails(self):
        formatter.create(self.spec, self.output)
        def mutate(core):
            node = ET.SubElement(core, "{http://purl.org/dc/elements/1.1/}creator")
            node.text = "Identifying author"
        self.rewrite_part("docProps/core.xml", mutate)
        self.assertTrue(formatter.check(self.output)["passed"])
        self.assertFalse(formatter.check(self.output, require_anonymous_metadata=True)["passed"])

    def test_existing_output_is_preserved(self):
        self.output.write_bytes(b"existing document")
        with self.assertRaises(FileExistsError):
            formatter.create(self.spec, self.output)
        self.assertEqual(self.output.read_bytes(), b"existing document")

    def test_unsupported_formal_mode_is_refused(self):
        spec = copy.deepcopy(self.spec)
        for mode in ("redhead", "minutes", "external_letter"):
            spec["mode"] = mode
            with self.assertRaisesRegex(ValueError, "native template"):
                formatter.create(spec, self.output)

    def test_missing_text_or_invalid_tables_are_refused(self):
        for blocks in (
                [{"role": "body"}],
                [{"role": "table", "rows": [["a", "b"], ["c"]]}],
                [{"role": "imprint", "text": "不适用于日常公文"}],
                [{"role": "body", "text": "禁止\x01控制字符"}]):
            with self.assertRaises((ValueError, TypeError)):
                formatter.create({"blocks": blocks}, self.output)

    def test_unrecognized_document_roles_are_not_certified(self):
        formatter.create(self.spec, self.output)
        def mutate(doc):
            pstyle = doc.find("w:body/w:p/w:pPr/w:pStyle", formatter.NS)
            pstyle.set(formatter.wtag("val"), "Normal")
        self.rewrite_part("word/document.xml", mutate)
        report = formatter.check(self.output)
        self.assertFalse(report["passed"])
        self.assertTrue(any("recognized role" in x for x in report["errors"]))


if __name__ == "__main__":
    unittest.main()
