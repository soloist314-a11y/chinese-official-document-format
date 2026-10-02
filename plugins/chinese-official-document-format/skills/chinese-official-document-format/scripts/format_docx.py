#!/usr/bin/env python3
"""Create and structurally check daily Chinese official documents.

Standard-library-only OOXML writer. It does not render pages or verify installed
fonts. It deliberately requires a native template for redhead documents,
meeting minutes, and external letters: first-page margins and even-page imprint
placement cannot be certified by changing a single section's margins.
"""

import argparse
import json
import sys
import unicodedata
from pathlib import Path
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W, "r": R}
ET.register_namespace("w", W)
ET.register_namespace("r", R)

PAGE = {"w": 11906, "h": 16838}
MARGINS = {"top": 2098, "bottom": 1871, "left": 1587, "right": 1474}
BODY_FONT = "仿宋_GB2312"
BODY_SIZE = 32  # Half-points: 三号 is 16 pt.
LINE = 580  # Twips: 29 pt.
GRID_LINE = round((PAGE["h"] - MARGINS["top"] - MARGINS["bottom"]) / 22)
# Additional character pitch uses 1/4096 pt, on top of the 16 pt base font.
CHAR_SPACE = round(((PAGE["w"] - MARGINS["left"] - MARGINS["right"])
                    / 20 / 28 - 16) * 4096)
GRID = {"type": "linesAndChars", "linePitch": GRID_LINE, "charSpace": CHAR_SPACE}
RULES = {
    "title": ("方正小标宋简体", 44, True, "center", 0, 640),
    "h1": ("黑体", 32, True, "both", 200, LINE),
    "h2": ("楷体_GB2312", 32, True, "both", 200, LINE),
    "h3": (BODY_FONT, 32, False, "both", 200, LINE),
    "h4": (BODY_FONT, 32, False, "both", 200, LINE),
    "h5": (BODY_FONT, 32, False, "both", 200, LINE),
    "body": (BODY_FONT, 32, False, "both", 200, LINE),
    "recipient": (BODY_FONT, 32, False, "both", 0, LINE),
    "attachment_item": (BODY_FONT, 32, False, "both", 200, LINE),
    "appendix_marker": ("黑体", 32, True, "both", 0, LINE),
    "signature": (BODY_FONT, 32, False, "right", 0, LINE),
    "date": (BODY_FONT, 32, False, "right", 0, LINE),
    "table_title": ("黑体", 32, True, "center", 0, LINE),
    "blank": (BODY_FONT, 32, False, "both", 0, LINE),
}
CONTENT_ROLES = set(RULES) - {"blank"}
SUPPORTED_ROLES = CONTENT_ROLES | {"blank", "table", "page_break", "attachment_list"}
LABEL_GROUP = {"attachment_item", "signature", "date"}


def wtag(name):
    return "{" + W + "}" + name


def sub(parent, name, attrs=None, text=None):
    node = ET.SubElement(parent, wtag(name))
    for key, value in (attrs or {}).items():
        node.set(wtag(key), str(value))
    if text is not None:
        node.text = text
    return node


def xml_bytes(root):
    # OOXML paragraph properties have a defined element order.
    paragraph_order = ("pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
                       "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
                       "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct",
                       "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
                       "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap",
                       "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl",
                       "divId", "cnfStyle", "rPr", "sectPr", "pPrChange")
    rank = {wtag(name): i for i, name in enumerate(paragraph_order)}
    for props in root.iter(wtag("pPr")):
        props[:] = sorted(props, key=lambda child: rank.get(child.tag, len(rank)))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def text_width_chars(text):
    """Estimate CJK character units; callers can provide exact indentation."""
    return sum(1 if unicodedata.east_asian_width(c) in "WF" else 0.5 for c in text)


def require_text(value, location):
    if not isinstance(value, str):
        raise ValueError(location + " must be a string; text is never rewritten.")
    if any(ord(c) < 32 and c not in "\n\r\t" for c in value):
        raise ValueError(location + " contains characters forbidden by XML 1.0.")
    return value


def normalize(spec):
    if not isinstance(spec, dict):
        raise ValueError("The input root must be a JSON object.")
    if spec.get("mode", "daily") != "daily":
        raise ValueError("Only daily mode can be created. Redhead/minutes/external "
                         "letters require a native template; do not simulate "
                         "first-page margins with blank paragraphs.")
    if not isinstance(spec.get("blocks"), list) or not spec["blocks"]:
        raise ValueError("blocks must be a nonempty list.")
    blocks = []
    for i, original in enumerate(spec["blocks"]):
        if not isinstance(original, dict):
            raise ValueError("blocks[%d] must be an object." % i)
        block = dict(original)
        role = block.get("role")
        if role not in SUPPORTED_ROLES:
            raise ValueError("Unsupported role %r in blocks[%d]." % (role, i))
        if role == "attachment_list":
            items = block.get("items")
            if not isinstance(items, list) or not items:
                raise ValueError("attachment_list.items must be nonempty.")
            for j, item in enumerate(items):
                item = {"text": item} if isinstance(item, str) else dict(item)
                item["role"] = "attachment_item"
                item["text"] = require_text(item.get("text"), "attachment item")
                if j == 0:
                    item["space_before_pt"] = 29
                blocks.append(item)
            continue
        if role == "table":
            rows = block.get("rows")
            if not isinstance(rows, list) or not rows or not rows[0]:
                raise ValueError("table.rows must be a nonempty rectangular list.")
            ncols = len(rows[0])
            for row in rows:
                if not isinstance(row, list) or len(row) != ncols:
                    raise ValueError("Every table row must have the same number of cells.")
                for cell in row:
                    require_text(cell, "table cell")
            size = block.get("font_size_pt", 12)
            line = block.get("line_pt", 16)
            headers = block.get("header_rows", 1)
            if size not in (10.5, 12) or not 14 <= line <= 18:
                raise ValueError("Table size must be 10.5 or 12 pt; line must be 14–18 pt.")
            if not isinstance(headers, int) or not 1 <= headers <= len(rows):
                raise ValueError("Tables require at least one repeating header row.")
        elif role in CONTENT_ROLES:
            block["text"] = require_text(block.get("text"), "blocks[%d].text" % i)
        blocks.append(block)
    return blocks


def add_text_run(paragraph, text, font, size, bold=False):
    run = sub(paragraph, "r")
    props = sub(run, "rPr")
    sub(props, "rFonts", {"ascii": font, "hAnsi": font, "eastAsia": font, "cs": font})
    sub(props, "b", {"val": "1" if bold else "0"})
    sub(props, "bCs", {"val": "1" if bold else "0"})
    sub(props, "color", {"val": "000000"})
    sub(props, "sz", {"val": size})
    sub(props, "szCs", {"val": size})
    sub(props, "lang", {"val": "zh-CN", "eastAsia": "zh-CN"})
    # Preserve newlines and tabs as OOXML breaks/tabs; never alter punctuation.
    parts = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    for i, line in enumerate(parts):
        if i:
            sub(run, "br")
        for j, segment in enumerate(line.split("\t")):
            if j:
                sub(run, "tab")
            t = sub(run, "t", text=segment)
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return run


def set_spacing(p, key, twips):
    pp = p.find("w:pPr", NS)
    if pp is None:
        pp = ET.Element(wtag("pPr"))
        p.insert(0, pp)
    spacing = pp.find("w:spacing", NS)
    if spacing is None:
        spacing = sub(pp, "spacing", {"line": LINE, "lineRule": "exact"})
    spacing.set(wtag(key), str(twips))


def set_keep_next(p):
    props = p.find("w:pPr", NS)
    if props is not None and props.find("w:keepNext", NS) is None:
        sub(props, "keepNext")


def add_paragraph(body, block):
    role = block["role"]
    font, size, bold, align, indent, line = RULES[role]
    p = sub(body, "p")
    props = sub(p, "pPr")
    sub(props, "pStyle", {"val": "Official_" + role})
    sub(props, "keepLines")
    sub(props, "widowControl")
    if role in {"title", "h1", "h2", "h3", "h4", "h5", "table_title"}:
        sub(props, "keepNext")
    if role == "appendix_marker":
        sub(props, "pageBreakBefore")
        sub(props, "keepNext")
    before = block.get("space_before_pt", 0)
    if role == "recipient":
        before = 29
    if role == "signature":
        before = 116  # Four body lines of space.
    sub(props, "spacing", {"before": round(before * 20), "after": 0,
                           "line": line, "lineRule": "exact"})
    if role in {"signature", "date"}:
        right_chars = block.get("right_indent_chars", 4)
        if role == "signature" and "right_indent_chars" not in block and block.get("align_with_date"):
            right_chars = max(0, 4 + (text_width_chars(block["align_with_date"])
                                      - text_width_chars(block["text"])) / 2)
        sub(props, "ind", {"firstLineChars": 0, "rightChars": round(right_chars * 100)})
    elif role == "attachment_item" and block.get("continuation_indent_chars") is not None:
        # First line starts at two characters; wrapping aligns with user-specified label width.
        continuation = block["continuation_indent_chars"]
        if not isinstance(continuation, (int, float)) or continuation < 2:
            raise ValueError("continuation_indent_chars must be at least 2.")
        sub(props, "ind", {"leftChars": round(continuation * 100),
                           "hangingChars": round((continuation - 2) * 100)})
    else:
        sub(props, "ind", {"firstLineChars": indent})
    sub(props, "jc", {"val": align})
    sub(props, "snapToGrid", {"val": "0"})
    add_text_run(p, block.get("text", ""), font, size, bold)
    return p


def add_table(body, block):
    tbl = sub(body, "tbl")
    props = sub(tbl, "tblPr")
    sub(props, "tblW", {"w": 0, "type": "auto"})
    sub(props, "jc", {"val": "center"})
    borders = sub(props, "tblBorders")
    for name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        sub(borders, name, {"val": "single", "sz": 4, "color": "000000"})
    sub(props, "tblLayout", {"type": "autofit"})
    cell_margins = sub(props, "tblCellMar")
    for name in ("top", "bottom", "left", "right"):
        sub(cell_margins, name, {"w": 80, "type": "dxa"})
    grid = sub(tbl, "tblGrid")
    width = PAGE["w"] - MARGINS["left"] - MARGINS["right"]
    columns = len(block["rows"][0])
    for column in range(columns):
        # Initial grid only: autoFit can rebalance widths after content layout.
        sub(grid, "gridCol", {"w": width // columns + (1 if column < width % columns else 0)})
    for i, row in enumerate(block["rows"]):
        tr = sub(tbl, "tr")
        if i < block.get("header_rows", 1):
            sub(sub(tr, "trPr"), "tblHeader")
        for text in row:
            tc = sub(tr, "tc")
            tcprops = sub(tc, "tcPr")
            sub(tcprops, "vAlign", {"val": "center"})
            p = sub(tc, "p")
            pp = sub(p, "pPr")
            sub(pp, "pStyle", {"val": "Official_table"})
            sub(pp, "spacing", {"before": 0, "after": 0,
                                "line": round(block.get("line_pt", 16) * 20), "lineRule": "exact"})
            sub(pp, "ind", {"firstLineChars": 0})
            sub(pp, "jc", {"val": "center"})
            sub(pp, "snapToGrid", {"val": "0"})
            add_text_run(p, text, "宋体", round(block.get("font_size_pt", 12) * 2))
    return tbl


def make_styles():
    root = ET.Element(wtag("styles"))
    defaults = sub(root, "docDefaults")
    rp = sub(sub(defaults, "rPrDefault"), "rPr")
    sub(rp, "rFonts", {"ascii": BODY_FONT, "hAnsi": BODY_FONT, "eastAsia": BODY_FONT})
    sub(rp, "sz", {"val": BODY_SIZE})
    pp = sub(sub(defaults, "pPrDefault"), "pPr")
    sub(pp, "spacing", {"line": LINE, "lineRule": "exact", "before": 0, "after": 0})
    for role in list(RULES) + ["table"]:
        style = sub(root, "style", {"type": "paragraph", "styleId": "Official_" + role})
        sub(style, "name", {"val": "Official " + role})
        sub(style, "qFormat")
    return root


def make_footer():
    footer = ET.Element(wtag("ftr"))
    p = sub(footer, "p")
    pp = sub(p, "pPr")
    sub(pp, "jc", {"val": "center"})
    # The attachment leaves daily page-number font/size to the application default.
    field = sub(p, "fldSimple", {"instr": " PAGE "})
    r = sub(field, "r")
    sub(r, "t", text="1")
    return footer


def relationship_root():
    return ET.Element("{" + PKG + "}Relationships")


def rel(root, rid, kind, target):
    ET.SubElement(root, "{" + PKG + "}Relationship", {
        "Id": rid, "Type": R + "/" + kind, "Target": target})


def create(spec, output):
    blocks = normalize(spec)
    document = ET.Element(wtag("document"))
    body = sub(document, "body")
    previous_p = None
    previous_role = None
    table_pending = False
    for block in blocks:
        role = block["role"]
        if role == "table":
            if previous_p is not None:
                set_spacing(previous_p, "after", 290)  # Half of one body line.
            add_table(body, block)
            previous_p = None
            table_pending = True
        elif role == "page_break":
            p = sub(body, "p")
            sub(sub(p, "r"), "br", {"type": "page"})
            previous_p = None
        else:
            if role in LABEL_GROUP and previous_p is not None:
                set_keep_next(previous_p)  # Bind to preceding content; visual QA remains necessary.
            p = add_paragraph(body, block)
            if table_pending:
                current_before = int(p.find("w:pPr/w:spacing", NS).get(wtag("before"), 0))
                set_spacing(p, "before", max(290, current_before))
                table_pending = False
            if role == "attachment_item" and previous_role != "attachment_item":
                set_spacing(p, "before", 580)
            previous_p = p
        previous_role = role
    sect = sub(body, "sectPr")
    sub(sect, "footerReference", {"type": "default"}).set("{" + R + "}id", "rIdFooter")
    sub(sect, "pgSz", PAGE)
    sub(sect, "pgMar", dict(MARGINS, header=720, footer=720, gutter=0))
    sub(sect, "cols", {"space": 720})
    sub(sect, "docGrid", GRID)
    settings = ET.Element(wtag("settings"))
    sub(settings, "updateFields", {"val": "true"})
    # Neutral metadata: no creator, company, source file, author, or device path.
    core = ET.Element("{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}coreProperties")
    app = ET.Element("{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Properties")
    docrels = relationship_root()
    rel(docrels, "rIdStyles", "styles", "styles.xml")
    rel(docrels, "rIdSettings", "settings", "settings.xml")
    rel(docrels, "rIdFooter", "footer", "footer1.xml")
    package_rels = relationship_root()
    rel(package_rels, "rIdDocument", "officeDocument", "word/document.xml")
    ET.SubElement(package_rels, "{" + PKG + "}Relationship", {
        "Id": "rIdCore", "Type": "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        "Target": "docProps/core.xml"})
    rel(package_rels, "rIdApp", "extended-properties", "docProps/app.xml")
    content = ET.Element("{" + CT + "}Types")
    ET.SubElement(content, "{" + CT + "}Default", {"Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"})
    ET.SubElement(content, "{" + CT + "}Default", {"Extension": "xml", "ContentType": "application/xml"})
    for name, kind in {
        "word/document.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "word/styles.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        "word/settings.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        "word/footer1.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        "docProps/core.xml": "application/vnd.openxmlformats-package.core-properties+xml",
        "docProps/app.xml": "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    }.items():
        ET.SubElement(content, "{" + CT + "}Override", {"PartName": "/" + name, "ContentType": kind})
    parts = {
        "[Content_Types].xml": content, "_rels/.rels": package_rels,
        "word/document.xml": document, "word/_rels/document.xml.rels": docrels,
        "word/styles.xml": make_styles(), "word/settings.xml": settings,
        "word/footer1.xml": make_footer(), "docProps/core.xml": core, "docProps/app.xml": app,
    }
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output, "x", ZIP_DEFLATED) as archive:
        for name, root in parts.items():
            archive.writestr(name, xml_bytes(root))
    report = check(output)
    if report["errors"]:
        raise ValueError("Generated file failed structural check: " + "; ".join(report["errors"]))
    return report


def att(node, name, default=None):
    return node.get(wtag(name), default) if node is not None else default


def check(path, require_anonymous_metadata=False):
    errors, warnings = [], []
    report = {"file": str(Path(path)), "profile": "daily", "errors": errors, "warnings": warnings,
              "validation": "OOXML structural validation only; not visual or pagination validation."}
    try:
        with ZipFile(path) as archive:
            bad = archive.testzip()
            if bad:
                errors.append("ZIP CRC error in " + bad)
            doc = ET.fromstring(archive.read("word/document.xml"))
            styles = ET.fromstring(archive.read("word/styles.xml"))
            settings = ET.fromstring(archive.read("word/settings.xml"))
            for name in archive.namelist():
                if name.startswith("docProps/") and name.endswith(".xml"):
                    metadata = ET.fromstring(archive.read(name))
                    for node in metadata.iter():
                        if node.tag.rsplit("}", 1)[-1] in {"creator", "lastModifiedBy", "Company", "Manager"} and (node.text or "").strip():
                            (errors if require_anonymous_metadata else warnings).append(
                                "Identifying metadata present: " + node.tag.rsplit("}", 1)[-1])
                if name.endswith(".rels"):
                    for node in ET.fromstring(archive.read(name)):
                        if node.get("TargetMode") == "External":
                            warnings.append("External relationship requires privacy review: " + name)
                if name in {"word/comments.xml", "word/people.xml"}:
                    (errors if require_anonymous_metadata else warnings).append(
                        "Review/people metadata present: " + name)
            if "word/numbering.xml" in archive.namelist():
                warnings.append("A numbering part exists; confirm no automatic numbering is used.")
    except (OSError, BadZipFile, KeyError, ET.ParseError) as exc:
        errors.append("Cannot read DOCX package: " + str(exc))
        report["passed"] = False
        return report
    for tree in (doc, styles):
        if tree.findall(".//w:numPr", NS):
            errors.append("Automatic numbering is present; supply numbering as literal text.")
    if doc.findall(".//w:ins", NS) or doc.findall(".//w:del", NS):
        (errors if require_anonymous_metadata else warnings).append(
            "Tracked revisions are present; review if a clean or anonymized copy is requested.")
    sections = doc.findall(".//w:sectPr", NS)
    if not sections:
        errors.append("No section properties were found; page settings cannot be verified.")
    for i, section in enumerate(sections, 1):
        margins = section.find("w:pgMar", NS)
        for key, expected in MARGINS.items():
            try:
                actual = int(att(margins, key, -1))
            except ValueError:
                actual = -1
            if abs(actual - expected) > 1:
                errors.append("Section %d %s margin differs from daily profile." % (i, key))
        size = section.find("w:pgSz", NS)
        for key, expected in PAGE.items():
            if att(size, key) != str(expected):
                errors.append("Section %d is not A4 portrait." % i)
        grid = section.find("w:docGrid", NS)
        if att(grid, "type") != "linesAndChars" or att(grid, "linePitch") != str(GRID_LINE):
            errors.append("Section %d grid does not divide the daily body area into 22 lines." % i)
        try:
            actual_chars = (PAGE["w"] - MARGINS["left"] - MARGINS["right"]) / 20 / (16 + int(att(grid, "charSpace", 0)) / 4096)
            if abs(actual_chars - 28) > .05:
                errors.append("Section %d character grid does not correspond to 28 characters." % i)
        except (ValueError, ZeroDivisionError):
            errors.append("Section %d character grid is invalid." % i)
        if section.find("w:titlePg", NS) is not None:
            warnings.append("A first-page header/footer variant exists; inspect its page numbering.")
    if settings.find("w:evenAndOddHeaders", NS) is not None:
        warnings.append("Daily profile does not require separate odd/even footers.")
    for i, p in enumerate(doc.findall(".//w:p", NS), 1):
        pp = p.find("w:pPr", NS)
        style = att(pp.find("w:pStyle", NS) if pp is not None else None, "val", "")
        role = style[len("Official_"):] if style.startswith("Official_") else style
        if role not in RULES and role != "table":
            if p.findall(".//w:t", NS):
                errors.append("Paragraph %d has no explicit recognized role; supply semantic role mapping before certification." % i)
            continue
        if role == "table":
            expected = ("宋体", None, False, "center", 0, None)
        else:
            expected = RULES[role]
        font, size, bold, align, indent, line = expected
        spacing = pp.find("w:spacing", NS)
        if att(spacing, "lineRule") != "exact":
            errors.append("Paragraph %d requires exact line spacing." % i)
        try:
            actual_line = int(att(spacing, "line", 0))
        except ValueError:
            actual_line = 0
        if line is not None and actual_line != line:
            errors.append("Paragraph %d line spacing differs from role %s." % (i, role))
        if role == "table" and not 280 <= actual_line <= 360:
            errors.append("Paragraph %d table line spacing must be 14–18 pt." % i)
        if att(pp.find("w:jc", NS), "val") != align:
            errors.append("Paragraph %d alignment differs from role %s." % (i, role))
        ind = pp.find("w:ind", NS)
        if role == "attachment_item" and att(ind, "hangingChars") is not None:
            try:
                if int(att(ind, "leftChars", 0)) - int(att(ind, "hangingChars", 0)) != 200:
                    errors.append("Paragraph %d attachment first line must start at two characters." % i)
            except ValueError:
                errors.append("Paragraph %d attachment indentation is invalid." % i)
        elif att(ind, "firstLineChars") != str(indent):
            errors.append("Paragraph %d first-line indentation differs from role %s." % (i, role))
        if role == "date" and att(ind, "rightChars") != "400":
            errors.append("Paragraph %d date must leave four characters on the right." % i)
        if role == "appendix_marker" and pp.find("w:pageBreakBefore", NS) is None:
            errors.append("Paragraph %d appendix must start on a new page." % i)
        for run in p.findall("w:r", NS):
            rp = run.find("w:rPr", NS)
            if rp is None:
                errors.append("Paragraph %d contains text without explicit font properties." % i)
                continue
            fonts = rp.find("w:rFonts", NS)
            if att(fonts, "eastAsia") != font:
                errors.append("Paragraph %d Chinese font differs from role %s." % (i, role))
            run_size = att(rp.find("w:sz", NS), "val")
            if size is not None and run_size != str(size):
                errors.append("Paragraph %d size differs from role %s." % (i, role))
            if role == "table" and run_size not in {"21", "24"}:
                errors.append("Paragraph %d table font must be 10.5 or 12 pt." % i)
            actual_bold = att(rp.find("w:b", NS), "val", "0") not in {"0", "false", "off"}
            if actual_bold != bold:
                errors.append("Paragraph %d bold setting differs from role %s." % (i, role))
    for i, table in enumerate(doc.findall(".//w:tbl", NS), 1):
        if not table.findall("w:tr/w:trPr/w:tblHeader", NS):
            errors.append("Table %d needs a repeating header row." % i)
        if table.findall(".//w:trHeight", NS):
            errors.append("Table %d has fixed row heights; allow content to determine heights." % i)
    warnings.extend([
        "Rendering is still required: inspect font substitution, clipping, page breaks, attachment/signature grouping and table continuation.",
        "The 28-character/22-line grid is encoded in OOXML; paragraph grid snapping is disabled to prioritize exact 29 pt spacing. Word/WPS must be checked visually.",
        "Text privacy is semantic: this check detects metadata, not organization names or personal details in the body.",
    ])
    report["passed"] = not errors
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    commands = parser.add_subparsers(dest="command", required=True)
    create_cmd = commands.add_parser("create", help="Create a daily-format DOCX with neutral metadata from UTF-8 JSON.")
    create_cmd.add_argument("input", type=Path, help="JSON with mode=daily and an explicit-role blocks list.")
    create_cmd.add_argument("output", type=Path, help="Destination .docx file.")
    create_cmd.add_argument("--report", type=Path, help="Write the structural check report as UTF-8 JSON.")
    check_cmd = commands.add_parser("check", help="Check a DOCX structurally against the daily profile.")
    check_cmd.add_argument("input", type=Path)
    check_cmd.add_argument("--report", type=Path)
    check_cmd.add_argument("--anonymous-metadata", action="store_true",
                           help="Treat identifying/review metadata as errors; does not anonymize body text.")
    args = parser.parse_args(argv)
    try:
        if args.command == "create":
            with args.input.open("r", encoding="utf-8-sig") as source:
                spec = json.load(source)
            report = create(spec, args.output)
        else:
            report = check(args.input, require_anonymous_metadata=args.anonymous_metadata)
        encoded = json.dumps(report, ensure_ascii=False, indent=2)
        if args.report:
            args.report.parent.mkdir(parents=True, exist_ok=True)
            args.report.write_text(encoded + "\n", encoding="utf-8")
        print(encoded)
        return 0 if report["passed"] else 1
    except (ValueError, OSError, TypeError) as exc:
        print("Error: " + str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
