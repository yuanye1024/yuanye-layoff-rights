#!/usr/bin/env python3
"""Build a three-part case report from a retained DOCX and reviewed JSON.

This script formats supplied content; it does not invent facts, perform legal
analysis, or calculate compensation. Requires Python 3.10+ and lxml; use the
host's bundled workspace Python when one is available.
"""

import argparse
from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
KINDS = {"body", "detail", "subheading", "law_heading", "quote", "formula", "result", "claim", "source_label", "source", "signature"}
SECTIONS = ["适用法律条例", "相关争议问题及条例", "本人诉求及相关条例依据"]


def element(name, **attributes):
    node = etree.Element(f"{{{W}}}{name}")
    for key, value in attributes.items():
        node.set(f"{{{W}}}{key}", str(value))
    return node


def encode(root):
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def validate_content(data):
    if not isinstance(data, dict) or not isinstance(data.get("title"), str) or not data["title"].strip():
        raise ValueError("A nonempty report title is required")
    sections = data.get("sections", [])
    if not isinstance(sections, list) or [s.get("title") for s in sections] != SECTIONS:
        raise ValueError("The three report sections must use the defined order and titles")
    for section in sections:
        if not isinstance(section.get("blocks"), list) or not section["blocks"]:
            raise ValueError("Every section requires reviewed content")
        previous_kind = None
        for index, block in enumerate(section["blocks"]):
            if block.get("kind") == "law_heading":
                basis = block.get("basis_for")
                if not isinstance(basis, str) or not basis.strip():
                    raise ValueError("Every legal citation must identify the specific issue, request or calculation it supports")
                if index + 1 >= len(section["blocks"]) or section["blocks"][index + 1].get("kind") != "quote":
                    raise ValueError("Every legal citation must be followed by its quoted text")
            if block.get("kind") == "quote" and previous_kind not in {"law_heading", "quote"}:
                raise ValueError("A legal quotation requires an adjacent citation with a support description")
            previous_kind = block.get("kind")
    for block in [b for s in sections for b in s["blocks"]] + data.get("closing", []):
        if not isinstance(block, dict) or block.get("kind") not in KINDS:
            raise ValueError("Unknown paragraph kind")
        if not isinstance(block.get("text"), str) or not block["text"].strip():
            raise ValueError("Empty or nontext paragraph")
        if block["text"].startswith(("现状：", "判断：", "建议：")):
            raise ValueError("A letter to the employer cannot contain chat-analysis labels")
        for spans_key in ("bold_spans", "highlight_spans"):
            if any(not isinstance(s, str) or not s or s not in block["text"] for s in block.get(spans_key, [])):
                raise ValueError("Every emphasis span must occur in its paragraph")


def build(template, content, output, font, overwrite=False):
    template, output = template.resolve(), output.resolve()
    if template == output:
        raise ValueError("The retained reference cannot be overwritten")
    if output.exists() and not overwrite:
        raise FileExistsError("Output exists; use --overwrite only for an authorized revision")
    validate_content(content)
    original_hash = hashlib.sha256(template.read_bytes()).hexdigest()
    with ZipFile(template) as source:
        # The source package is the working-copy authority, including styles,
        # theme, numbering, settings and the exact section geometry.
        package = {i.filename: source.read(i.filename) for i in source.infolist() if not i.is_dir()}
    root = etree.fromstring(package["word/document.xml"])
    body = root.find("w:body", NS)
    if body is None or len(body.xpath(".//w:sectPr", namespaces=NS)) != 1:
        raise ValueError("This builder supports a verified single-section reference")
    section = body.find("w:sectPr", NS)
    if section is None:
        raise ValueError("A final body section property is required")
    section = deepcopy(section)
    source_paragraphs = body.findall("w:p", NS)
    anchors = {0:"关于解除劳动合同",1:SECTIONS[0],8:SECTIONS[1],30:SECTIONS[2],62:"提出人"}
    for index, prefix in anchors.items():
        if len(source_paragraphs) <= index or not "".join(source_paragraphs[index].xpath('.//w:t/text()', namespaces=NS)).startswith(prefix):
            raise ValueError("Reference paragraph roles changed; inspect and update the template contract")
    prototypes = {"title":0,"body":22,"detail":10,"subheading":9,"law_heading":14,"quote":14,"formula":10,"result":50,"claim":31,"source_label":66,"source":67,"signature":62}
    # These citation styles implement the user's explicit typography change.
    # All existing source styles remain intact; no case facts or legal rules
    # are part of this formatter.
    styles = etree.fromstring(package["word/styles.xml"])
    for style_id, title, bold, before, after, keep_next in [
        ("LegalCitation", "法律依据导语", True, 100, 40, True),
        ("LegalQuote", "法律条文引用", False, 0, 100, False),
    ]:
        style = element("style", type="paragraph", styleId=style_id, customStyle=1)
        style.append(element("name", val=title))
        style.append(element("basedOn", val="1"))
        props = element("pPr")
        if keep_next: props.append(element("keepNext"))
        props.append(element("widowControl"))
        props.append(element("snapToGrid", val=0))
        props.append(element("shd", val="clear", color="auto", fill="EDEDED"))
        props.append(element("spacing", before=before, after=after, line=300, lineRule="auto"))
        props.append(element("ind", left=420, right=420, firstLine=0, firstLineChars=0))
        props.append(element("jc", val="both"))
        style.append(props)
        run_props = element("rPr")
        run_props.append(element("rFonts", eastAsia="宋体"))
        if bold:
            run_props.append(element("b")); run_props.append(element("bCs"))
        if style_id == "LegalQuote":
            run_props.append(element("i")); run_props.append(element("iCs"))
        run_props.append(element("color", val="404040"))
        run_props.append(element("sz", val=19)); run_props.append(element("szCs", val=19))
        style.append(run_props)
        styles.append(style)
    package["word/styles.xml"] = encode(styles)
    highlight = deepcopy(source_paragraphs[40].xpath("w:r/w:rPr/w:highlight[@w:val='yellow']", namespaces=NS)[0])
    links = []
    for node in list(body):
        body.remove(node)

    def paragraph(text, kind="body", source_index=None, bold_spans=None, highlight_spans=None, basis_for=None):
        prototype = source_paragraphs[source_index if source_index is not None else prototypes[kind]]
        p = element("p")
        pp = prototype.find("w:pPr", NS)
        if kind in {"law_heading", "quote"}:
            pp = element("pPr")
            pp.append(element("pStyle", val="LegalCitation" if kind == "law_heading" else "LegalQuote"))
            p.append(pp)
            if kind == "law_heading": text = f"法律依据（支持：{basis_for}）：{text}"
        elif pp is not None: p.append(deepcopy(pp))
        if kind in {"source_label", "subheading"}:
            pp = p.find("w:pPr", NS)
            if pp is None: pp = element("pPr"); p.insert(0, pp)
            keep = pp.find("w:keepNext", NS)
            if keep is None: keep = element("keepNext"); pp.insert(0, keep)
            keep.attrib.pop(f"{{{W}}}val", None)
        textual_runs = prototype.xpath(".//w:r[w:t]", namespaces=NS)
        rp = textual_runs[0].find("w:rPr", NS) if textual_runs else None
        def spans_to_ranges(spans):
            ranges = []
            for span in spans or []:
                start = 0
                while (start := text.find(span, start)) >= 0:
                    ranges.append((start, start + len(span)))
                    start += len(span)
            return ranges
        ranges = spans_to_ranges(bold_spans)
        highlights = spans_to_ranges(highlight_spans)
        cuts=sorted({0,len(text),*(x for pair in ranges + highlights for x in pair)})
        container = p
        if kind == "source" and text.startswith(("https://", "http://")):
            container = element("hyperlink")
            rel_id = f"rIdLetterLink{len(links) + 1}"
            container.set("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", rel_id)
            links.append((rel_id, text))
            p.append(container)
        for left,right in zip(cuts,cuts[1:]):
            run=element("r"); props=deepcopy(rp) if rp is not None else element("rPr")
            if kind == "quote":
                for name in ("i", "iCs"):
                    node = props.find(f"w:{name}", NS)
                    if node is None: node = element(name); props.append(node)
                    node.attrib.pop(f"{{{W}}}val", None)
            if font:
                fonts=props.find("w:rFonts",NS)
                if fonts is None: fonts=element("rFonts"); props.insert(0,fonts)
                fonts.set(f"{{{W}}}eastAsia",font)
                fonts.attrib.pop(f"{{{W}}}eastAsiaTheme",None)
            if any(a<=left and right<=b for a,b in ranges):
                for name in ("b","bCs"):
                    node=props.find(f"w:{name}",NS)
                    if node is None:node=element(name);props.append(node)
                    node.attrib.pop(f"{{{W}}}val",None)
            if any(a <= left and right <= b for a, b in highlights):
                existing = props.find("w:highlight", NS)
                if existing is not None: props.remove(existing)
                props.append(deepcopy(highlight))
            run.append(props)
            t=element("t");t.set("{http://www.w3.org/XML/1998/namespace}space","preserve");t.text=text[left:right]
            run.append(t);container.append(run)
        body.append(p)

    paragraph(content["title"], "title")
    for index,part in enumerate(content["sections"]):
        paragraph(part["title"], source_index=[1,8,30][index])
        for block in part["blocks"]:
            paragraph(block["text"], block["kind"], bold_spans=block.get("bold_spans"), highlight_spans=block.get("highlight_spans"), basis_for=block.get("basis_for"))
    for block in content.get("closing", []):
        paragraph(block["text"], block["kind"], source_index=63 if block["kind"]=="signature" and block["text"].startswith("时间") else None, bold_spans=block.get("bold_spans"), highlight_spans=block.get("highlight_spans"))
    body.append(section)
    package["word/document.xml"] = encode(root)

    rels_path = "word/_rels/document.xml.rels"
    rels = etree.fromstring(package[rels_path])
    for rel in list(rels):
        if rel.get("Type", "").endswith(("/image", "/hyperlink")):
            rels.remove(rel)
    rel_namespace = "http://schemas.openxmlformats.org/package/2006/relationships"
    for rel_id, url in links:
        etree.SubElement(rels, f"{{{rel_namespace}}}Relationship", Id=rel_id,
                         Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                         Target=url, TargetMode="External")
    package[rels_path] = encode(rels)
    # The sample's evidence images and personal metadata are explicitly outside
    # the output slot contract. Remove them from the generated working copy.
    removed = [name for name in package if name.startswith("word/media/") or name == "docProps/custom.xml"]
    for name in removed:
        package.pop(name)
    rels = etree.fromstring(package["_rels/.rels"])
    for rel in list(rels):
        if rel.get("Type", "").endswith("/custom-properties"):
            rels.remove(rel)
    package["_rels/.rels"] = encode(rels)
    types = etree.fromstring(package["[Content_Types].xml"])
    for node in list(types):
        if node.get("PartName", "").lstrip("/") in removed:
            types.remove(node)
    package["[Content_Types].xml"] = encode(types)
    core_ns = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    core = etree.Element(f"{{{core_ns}}}coreProperties", nsmap={"cp":core_ns,"dc":"http://purl.org/dc/elements/1.1/","dcterms":"http://purl.org/dc/terms/","xsi":"http://www.w3.org/2001/XMLSchema-instance"})
    etree.SubElement(core,"{http://purl.org/dc/elements/1.1/}title").text = content["title"]
    etree.SubElement(core,"{http://purl.org/dc/elements/1.1/}creator").text = "劳动争议助手"
    stamp = etree.SubElement(core,"{http://purl.org/dc/terms/}created")
    stamp.set("{http://www.w3.org/2001/XMLSchema-instance}type", "dcterms:W3CDTF")
    stamp.text = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    package["docProps/core.xml"] = encode(core)
    app = etree.fromstring(package["docProps/app.xml"])
    for node in app:
        if node.tag.split("}")[-1] in {"Company", "Manager", "Pages", "Words", "Characters", "CharactersWithSpaces"}:
            node.text = "" if node.tag.split("}")[-1] in {"Company", "Manager"} else "0"
    package["docProps/app.xml"] = encode(app)
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as target:
        for name, data in package.items():
            target.writestr(name, data)
    if hashlib.sha256(template.read_bytes()).hexdigest() != original_hash:
        raise RuntimeError("Retained reference changed during generation")
    print(json.dumps({"output":str(output),"paragraphs":len(body)-1,"reference_sha256":original_hash},ensure_ascii=False))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template", type=Path, default=Path(__file__).resolve().parents[1] / "assets" / "report-template.docx", help="Read-only format reference; defaults to the bundled neutral template")
    parser.add_argument("--content", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--font", default=None, help="Explicitly authorized CJK font replacement; omitted to retain the reference theme")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()
    build(args.template, json.loads(args.content.read_text(encoding="utf-8")), args.output, args.font, args.overwrite)
