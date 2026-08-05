from collections import Counter
import hashlib
import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.enum.section import WD_ORIENT
from lxml import etree


CAPTION_RE = re.compile(
    r"^(figure|fig\.)\s+\d+(?:\.\d+)*(?::|\b)",
    re.IGNORECASE,
)


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def validate(source_docx, output_docx):

    source_snapshot = _snapshot(
        source_docx
    )
    output_snapshot = _snapshot(
        output_docx
    )

    checks = [
        "image_count",
        "image_format_distribution",
        "image_relationship_count",
        "section_count",
        "landscape_section_count",
        "orientation_sequence",
        "figure_count",
        "caption_count",
        "caption_image_adjacency",
        "content_order_integrity",
    ]

    report = {}
    failed = 0

    print("\n========== VALIDATION REPORT ==========\n")
    print("CRITICAL CHECKS")
    print("----------------")

    for item in checks:

        source_value = source_snapshot[item]
        output_value = output_snapshot[item]

        passed = source_value == output_value

        report[item] = {
            "source": source_value,
            "output": output_value,
            "pass": passed,
            "severity": "CRITICAL",
        }

        if passed:
            print(
                f"[PASS] {item:<28} "
                f"Source={source_value} "
                f"Output={output_value}"
            )
        else:
            failed += 1
            print(
                f"[FAIL] {item:<28} "
                f"Source={source_value} "
                f"Output={output_value}"
            )

    overall_status = "PASS"

    if failed > 0:
        overall_status = "FAIL"

    print("\n======================================")
    print(f"Overall Status : {overall_status}")
    print(f"Critical Failed: {failed}")
    print("======================================\n")

    report["summary"] = {
        "status": overall_status,
        "critical_failed": failed,
    }

    return report


def _snapshot(docx_path):

    document = Document(
        docx_path
    )

    with ZipFile(docx_path) as archive:

        root = etree.fromstring(
            archive.read(
                "word/document.xml"
            )
        )

        rid_to_target = _extract_relationships(
            archive
        )

    body = root.xpath(
        "./w:body",
        namespaces=NS,
    )[0]

    top_level_items = []
    order_index = 0

    for element in body:

        if etree.QName(
            element
        ).localname == "sectPr":
            continue

        top_level_items.append(
            _classify_top_level_item(
                element,
                rid_to_target,
                order_index,
            )
        )
        order_index += 1

    grouped_items = _group_figure_blocks(
        top_level_items
    )

    image_refs = _collect_all_image_refs(
        body,
        rid_to_target,
    )

    format_distribution = Counter(
        image_ref["image_format"]
        for image_ref in image_refs
    )

    orientation_sequence = []

    for section in document.sections:

        if section.orientation == WD_ORIENT.LANDSCAPE:
            orientation_sequence.append(
                "landscape"
            )
        else:
            orientation_sequence.append(
                "portrait"
            )

    content_signature = [
        _content_signature(item)
        for item in grouped_items
    ]

    return {
        "image_count": len(image_refs),
        "image_format_distribution": dict(
            sorted(
                format_distribution.items()
            )
        ),
        "image_relationship_count": len(
            {
                image_ref["relationship_id"]
                for image_ref in image_refs
            }
        ),
        "section_count": len(
            document.sections
        ),
        "landscape_section_count": sum(
            1
            for orientation in orientation_sequence
            if orientation == "landscape"
        ),
        "orientation_sequence": orientation_sequence,
        "figure_count": sum(
            1
            for item in grouped_items
            if item["type"] == "figure"
        ),
        "caption_count": sum(
            1
            for item in top_level_items
            if item["type"] == "caption"
        ),
        "caption_image_adjacency": sum(
            1
            for item in grouped_items
            if item["type"] == "figure"
        ),
        "content_order_integrity": hashlib.sha1(
            json.dumps(
                content_signature,
                separators=(",", ":"),
            ).encode("utf-8")
        ).hexdigest(),
    }


def _extract_relationships(archive):

    rid_to_target = {}
    rels_path = "word/_rels/document.xml.rels"

    if rels_path not in archive.namelist():
        return rid_to_target

    rels_root = ET.fromstring(
        archive.read(rels_path)
    )

    for rel in rels_root:

        r_id = rel.get("Id")
        target = rel.get(
            "Target",
            "",
        )

        if r_id and target:
            rid_to_target[r_id] = target

    return rid_to_target


def _classify_top_level_item(
    element,
    rid_to_target,
    order_index,
):

    local_name = etree.QName(
        element
    ).localname

    if local_name == "tbl":
        return {
            "type": "table",
            "order_index": order_index,
        }

    if local_name != "p":
        return {
            "type": local_name,
            "order_index": order_index,
        }

    image_refs = _collect_image_refs_from_element(
        element,
        rid_to_target,
        order_index,
    )

    if image_refs:
        return {
            "type": "image",
            "order_index": order_index,
            "image_refs": image_refs,
        }

    textbox_texts = _get_textbox_texts(
        element
    )
    text = _get_visible_paragraph_text(
        element
    ).strip()

    if CAPTION_RE.match(text):
        return {
            "type": "caption",
            "order_index": order_index,
            "text": text,
        }

    if textbox_texts and not text:
        return {
            "type": "textbox",
            "order_index": order_index,
            "textbox_count": len(textbox_texts),
        }

    return {
        "type": "paragraph",
        "order_index": order_index,
        "text": text,
    }


def _group_figure_blocks(items):

    grouped_items = []
    index = 0

    while index < len(items):

        item = items[index]

        if item["type"] == "caption":

            images = []
            look_ahead = index + 1

            while (
                look_ahead < len(items)
                and items[look_ahead]["type"] == "image"
            ):
                images.append(
                    items[look_ahead]
                )
                look_ahead += 1

            if images:
                grouped_items.append(
                    {
                        "type": "figure",
                        "order_index": item["order_index"],
                        "caption_position": "before",
                        "caption": item,
                        "images": images,
                    }
                )
                index = look_ahead
                continue

        if item["type"] == "image":

            images = [item]
            look_ahead = index + 1

            while (
                look_ahead < len(items)
                and items[look_ahead]["type"] == "image"
            ):
                images.append(
                    items[look_ahead]
                )
                look_ahead += 1

            if (
                look_ahead < len(items)
                and items[look_ahead]["type"] == "caption"
            ):
                grouped_items.append(
                    {
                        "type": "figure",
                        "order_index": item["order_index"],
                        "caption_position": "after",
                        "caption": items[look_ahead],
                        "images": images,
                    }
                )
                index = look_ahead + 1
                continue

        grouped_items.append(item)
        index += 1

    return grouped_items


def _collect_all_image_refs(
    body,
    rid_to_target,
):

    image_refs = []

    for r_id in body.xpath(
        ".//*[local-name()='blip']/@r:embed",
        namespaces=NS,
    ):
        image_refs.append(
            _build_image_ref(
                r_id,
                rid_to_target,
                -1,
            )
        )

    for r_id in body.xpath(
        ".//*[local-name()='imagedata']/@r:id",
        namespaces=NS,
    ):
        image_refs.append(
            _build_image_ref(
                r_id,
                rid_to_target,
                -1,
            )
        )

    return image_refs


def _collect_image_refs_from_element(
    element,
    rid_to_target,
    order_index,
):

    image_refs = []
    seen = set()

    for r_id in element.xpath(
        ".//*[local-name()='blip']/@r:embed",
        namespaces=NS,
    ):

        if r_id in seen:
            continue

        seen.add(r_id)
        image_refs.append(
            _build_image_ref(
                r_id,
                rid_to_target,
                order_index,
            )
        )

    for r_id in element.xpath(
        ".//*[local-name()='imagedata']/@r:id",
        namespaces=NS,
    ):

        if r_id in seen:
            continue

        seen.add(r_id)
        image_refs.append(
            _build_image_ref(
                r_id,
                rid_to_target,
                order_index,
            )
        )

    return image_refs


def _build_image_ref(
    r_id,
    rid_to_target,
    order_index,
):

    target = rid_to_target.get(
        r_id,
        "",
    )

    return {
        "name": Path(target).name,
        "relationship_id": r_id,
        "image_format": Path(target).suffix.lower().lstrip("."),
        "document_position_index": order_index,
        "target": target,
    }


def _get_textbox_texts(element):

    textbox_texts = []
    altcontent_seen = set()

    for node in element.xpath(
        ".//*[local-name()='txbxContent']"
    ):

        altcontent = node.xpath(
            "ancestor::*[local-name()='AlternateContent'][1]"
        )

        if altcontent:

            alt_key = id(
                altcontent[0]
            )

            if alt_key in altcontent_seen:
                continue

            altcontent_seen.add(
                alt_key
            )

        text = "".join(
            node.xpath(
                ".//*[local-name()='t']/text()"
            )
        ).strip()

        if text:
            textbox_texts.append(
                text
            )

    return textbox_texts


def _get_visible_paragraph_text(element):

    text_parts = []

    for t_node in element.xpath(
        ".//*[local-name()='t']"
    ):

        if not t_node.text:
            continue

        ancestor = t_node.getparent()
        skip_node = False

        while ancestor is not None and ancestor is not element:

            local_name = etree.QName(
                ancestor
            ).localname

            if local_name in {
                "drawing",
                "txbxContent",
            }:
                skip_node = True
                break

            ancestor = ancestor.getparent()

        if not skip_node:
            text_parts.append(
                t_node.text
            )

    return "".join(
        text_parts
    )


def _content_signature(item):

    if item["type"] == "figure":
        image_count = sum(
            len(image["image_refs"])
            for image in item["images"]
        )

        return {
            "type": "figure",
            "caption_position": item["caption_position"],
            "image_count": image_count,
        }

    if item["type"] == "image":
        return {
            "type": "image",
            "image_count": len(
                item["image_refs"]
            ),
        }

    if item["type"] == "textbox":
        return {
            "type": "textbox",
            "textbox_count": item.get(
                "textbox_count",
                0,
            ),
        }

    return {
        "type": item["type"],
    }


if __name__ == "__main__":

    import sys

    if len(sys.argv) != 3:

        print(
            "Usage:\n"
            "python validator.py source.docx output.docx"
        )

        sys.exit(1)

    result = validate(
        sys.argv[1],
        sys.argv[2],
    )

    print(
        json.dumps(
            result,
            indent=4,
        )
    )