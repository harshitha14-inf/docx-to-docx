import json
import re
from copy import deepcopy
from datetime import date
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from models import FigureBlock
from style_engine.style_mapper import StyleMapper
from template_manager import TemplateManager


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
XML_NS = "{http://www.w3.org/XML/1998/namespace}"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_TYPE_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
REL_TYPE_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"

REL_ID_ATTR_LOCALNAMES = {"id", "embed", "link", "dm", "lo", "qs", "cs", "href"}

REL_TYPE_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
REL_TYPE_STYLES_WITH_EFFECTS = "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects"
REL_TYPE_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"

TEMPLATE_OWNED_REL_TYPES = {
    REL_TYPE_HEADER,
    REL_TYPE_FOOTER,
    REL_TYPE_STYLES,
    REL_TYPE_STYLES_WITH_EFFECTS,
    REL_TYPE_THEME,
}

CAPTION_NUMBER_RE = re.compile(r"^(figure|table)\s+\d+(\.\d+)*\s*", re.IGNORECASE)

HEADING_STYLE_IDS = {"Heading1", "Heading2", "Heading3", "Heading4"}

CANDIDATE_TABLE_STYLE_KEYS = [
    "infineon standard",
    "infineonstandard",
    "table grid",
    "tablegrid",
]


class Builder:

    def __init__(
        self,
        project_root=None,
        template_manager=None,
        style_mapper=None,
        formatting_policy=None,
    ):

        self.project_root = Path(project_root or Path(__file__).resolve().parent)
        self.template_manager = template_manager or TemplateManager(self.project_root)
        self.style_mapper = style_mapper
        self.formatting_policy = formatting_policy or self._load_formatting_policy()

    def build(
        self,
        model,
        out_file,
        document_type=None,
        selected_template=None,
        preserve_template_cover=True,
        apply_style_mapping=True,
        cover_metadata=None,
    ):

        if not model.source_docx_path:
            raise ValueError("DocumentModel.source_docx_path is required")

        base_docx_path, template_used = self._resolve_base_document(
            model,
            document_type=document_type,
            selected_template=selected_template,
        )

        active_style_mapper = self._resolve_style_mapper(
            apply_style_mapping=apply_style_mapping,
        )

        style_alias_map = self._load_template_style_aliases(base_docx_path)

        id_remap = {}
        merged_rels_bytes = None
        num_id_remap = {}
        merged_numbering_bytes = None
        settings_bytes = None

        if template_used:
            merged_rels_bytes, id_remap = self._compute_merged_relationships(
                base_docx_path=base_docx_path,
                source_docx_path=model.source_docx_path,
            )
            merged_numbering_bytes, num_id_remap = self._compute_merged_numbering(
                base_docx_path=base_docx_path,
                source_docx_path=model.source_docx_path,
            )
            settings_bytes = self._build_settings_with_update_fields(base_docx_path)

        resolved_cover_metadata = self._resolve_cover_metadata(
            model,
            overrides=cover_metadata,
        )

        first_heading_text = self._find_first_heading_text(model)

        field_resolver = self._make_field_resolver(
            resolved_cover_metadata,
            first_heading_text,
        )

        custom_props_bytes = None
        core_props_bytes = None

        if template_used:
            # Word recalculates DOCPROPERTY fields from docProps/custom.xml and
            # docProps/core.xml on open (now that updateFields=true is set) - the
            # underlying properties must be patched too, or Word will revert our
            # cached field text back to the template's placeholder values
            # (e.g. ConfidentialityMarking="restricted", Title="Document title").
            custom_props_bytes = self._build_custom_properties_xml(
                base_docx_path,
                resolved_cover_metadata,
            )
            core_props_bytes = self._build_core_properties_xml(
                model.source_docx_path,
                resolved_cover_metadata.get("Title"),
            )

        cover_content = self._extract_cover_content(model)

        content_width_twips = self._get_template_content_width_twips(base_docx_path)

        rebuilt_document = self._rebuild_document_xml(
            model,
            base_docx_path=base_docx_path,
            preserve_template_cover=preserve_template_cover and template_used,
            style_mapper=active_style_mapper,
            style_alias_map=style_alias_map,
            id_remap=id_remap,
            num_id_remap=num_id_remap,
            cover_content=cover_content,
            content_width_twips=content_width_twips,
        )

        rebuilt_document = self._update_fields_in_xml(rebuilt_document, field_resolver)

        with ZipFile(base_docx_path, "r") as source_zip:
            with ZipFile(model.source_docx_path, "r") as original_zip:
                with ZipFile(out_file, "w") as output_zip:

                    written_files = set()

                    for info in source_zip.infolist():

                        data = source_zip.read(info.filename)

                        if (
                            template_used
                            and info.filename == "[Content_Types].xml"
                            and info.filename in original_zip.namelist()
                        ):
                            data = self._merge_content_types(
                                template_types_xml=data,
                                source_types_xml=original_zip.read(info.filename),
                            )

                        elif template_used and info.filename == "word/_rels/document.xml.rels":
                            data = merged_rels_bytes

                        elif (
                            template_used
                            and info.filename == "word/numbering.xml"
                            and merged_numbering_bytes is not None
                        ):
                            data = merged_numbering_bytes

                        elif (
                            template_used
                            and info.filename == "word/settings.xml"
                            and settings_bytes is not None
                        ):
                            data = settings_bytes

                        elif (
                            template_used
                            and info.filename == "docProps/custom.xml"
                            and custom_props_bytes is not None
                        ):
                            data = custom_props_bytes

                        elif (
                            template_used
                            and info.filename == "docProps/core.xml"
                            and core_props_bytes is not None
                        ):
                            data = core_props_bytes

                        elif (
                            template_used
                            and self._should_overlay_from_source(info.filename)
                            and info.filename in original_zip.namelist()
                        ):
                            data = original_zip.read(info.filename)

                        elif info.filename == "word/document.xml":
                            data = rebuilt_document

                        elif template_used and self._is_header_or_footer_part(info.filename):
                            data = self._update_fields_in_xml(data, field_resolver)

                        output_zip.writestr(info, data)
                        written_files.add(info.filename)

                    if template_used:

                        for filename in original_zip.namelist():

                            if not self._should_overlay_from_source(filename):
                                continue

                            if filename in written_files:
                                continue

                            source_info = original_zip.getinfo(filename)

                            output_zip.writestr(
                                source_info,
                                original_zip.read(filename),
                            )
                            written_files.add(filename)

    # ------------------------------------------------------------------
    # Base document / style resolution
    # ------------------------------------------------------------------

    def _resolve_base_document(
        self,
        model,
        document_type,
        selected_template,
    ):

        if selected_template:
            template_path = Path(selected_template)

            if not template_path.exists():
                raise FileNotFoundError(f"Selected template not found: {template_path}")

            return str(template_path), True

        if document_type is not None:
            return self.template_manager.get_template(document_type), True

        return model.source_docx_path, False

    def _resolve_style_mapper(
        self,
        apply_style_mapping,
    ):

        if not apply_style_mapping:
            return None

        if self.style_mapper is not None:
            return self.style_mapper

        return StyleMapper.from_file(self.project_root / "config" / "style_map.json")

    # ------------------------------------------------------------------
    # Relationship-ID collision-safe merge
    # ------------------------------------------------------------------

    def _compute_merged_relationships(self, base_docx_path, source_docx_path):

        rels_part = "word/_rels/document.xml.rels"

        with ZipFile(base_docx_path, "r") as archive:
            template_rels_xml = archive.read(rels_part)

        with ZipFile(source_docx_path, "r") as archive:
            if rels_part not in archive.namelist():
                return template_rels_xml, {}
            source_rels_xml = archive.read(rels_part)

        template_root = etree.fromstring(template_rels_xml)
        source_root = etree.fromstring(source_rels_xml)

        rel_by_id = {}

        for rel in template_root:
            rel_id = rel.get("Id")
            if rel_id:
                rel_by_id[rel_id] = deepcopy(rel)

        used_numbers = set()

        for rel_id in rel_by_id:
            match = re.match(r"rId(\d+)$", rel_id)
            if match:
                used_numbers.add(int(match.group(1)))

        for rel in source_root:
            match = re.match(r"rId(\d+)$", rel.get("Id", ""))
            if match:
                used_numbers.add(int(match.group(1)))

        next_number = (max(used_numbers) + 1) if used_numbers else 1

        id_remap = {}

        for rel in source_root:

            rel_type = rel.get("Type", "")
            old_id = rel.get("Id")

            if not old_id:
                continue

            if rel_type in TEMPLATE_OWNED_REL_TYPES:
                # Migrated content never carries its own header/footer/styles/theme parts.
                continue

            new_rel = deepcopy(rel)

            if old_id in rel_by_id:
                new_id = f"rId{next_number}"
                next_number += 1
                new_rel.set("Id", new_id)
                id_remap[old_id] = new_id
            else:
                id_remap[old_id] = old_id

            rel_by_id[id_remap[old_id]] = new_rel

        merged_root = etree.Element(f"{{{REL_NS}}}Relationships")

        for rel_id in sorted(rel_by_id.keys(), key=lambda v: (len(v), v)):
            merged_root.append(rel_by_id[rel_id])

        merged_xml = etree.tostring(
            merged_root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

        return merged_xml, id_remap

    def _remap_relationship_ids(self, element, id_remap):

        if not id_remap:
            return

        for node in element.iter():
            for attr_name in list(node.attrib.keys()):
                if not attr_name.startswith(R_NS):
                    continue

                local_name = attr_name[len(R_NS):]

                if local_name not in REL_ID_ATTR_LOCALNAMES:
                    continue

                value = node.attrib[attr_name]
                mapped = id_remap.get(value)

                if mapped and mapped != value:
                    node.attrib[attr_name] = mapped

    # ------------------------------------------------------------------
    # Numbering-ID collision-safe merge (heading chapter numbers + lists)
    # ------------------------------------------------------------------

    def _compute_merged_numbering(self, base_docx_path, source_docx_path):

        numbering_part = "word/numbering.xml"

        with ZipFile(base_docx_path, "r") as archive:
            if numbering_part not in archive.namelist():
                return None, {}
            template_numbering_xml = archive.read(numbering_part)

        with ZipFile(source_docx_path, "r") as archive:
            if numbering_part not in archive.namelist():
                return template_numbering_xml, {}
            source_numbering_xml = archive.read(numbering_part)

        template_root = etree.fromstring(template_numbering_xml)
        source_root = etree.fromstring(source_numbering_xml)

        other_nodes = []
        abstract_by_id = {}
        num_by_id = {}
        ordered_abstract_ids = []
        ordered_num_ids = []

        for node in template_root:
            local_name = etree.QName(node).localname

            if local_name == "abstractNum":
                abstract_id = node.get(f"{W_NS}abstractNumId")
                abstract_by_id[abstract_id] = deepcopy(node)
                ordered_abstract_ids.append(abstract_id)
            elif local_name == "num":
                num_id = node.get(f"{W_NS}numId")
                num_by_id[num_id] = deepcopy(node)
                ordered_num_ids.append(num_id)
            else:
                # e.g. w:numPicBullet - preserve template-only, rare to need source's own.
                other_nodes.append(deepcopy(node))

        used_abstract_numbers = set()
        used_num_numbers = set()

        for node in list(template_root) + list(source_root):
            local_name = etree.QName(node).localname

            if local_name == "abstractNum":
                try:
                    used_abstract_numbers.add(int(node.get(f"{W_NS}abstractNumId")))
                except (TypeError, ValueError):
                    pass
            elif local_name == "num":
                try:
                    used_num_numbers.add(int(node.get(f"{W_NS}numId")))
                except (TypeError, ValueError):
                    pass

        next_abstract_number = (max(used_abstract_numbers) + 1) if used_abstract_numbers else 0
        next_num_number = (max(used_num_numbers) + 1) if used_num_numbers else 1

        abstract_id_remap = {}
        num_id_remap = {}

        for node in source_root:

            if etree.QName(node).localname != "abstractNum":
                continue

            old_id = node.get(f"{W_NS}abstractNumId")

            if old_id is None:
                continue

            new_node = deepcopy(node)

            if old_id in abstract_by_id:
                new_id = str(next_abstract_number)
                next_abstract_number += 1
                new_node.set(f"{W_NS}abstractNumId", new_id)
                abstract_id_remap[old_id] = new_id
            else:
                abstract_id_remap[old_id] = old_id

            abstract_by_id[abstract_id_remap[old_id]] = new_node
            ordered_abstract_ids.append(abstract_id_remap[old_id])

        for node in source_root:

            if etree.QName(node).localname != "num":
                continue

            old_id = node.get(f"{W_NS}numId")

            if old_id is None:
                continue

            new_node = deepcopy(node)

            abstract_ref = new_node.find("w:abstractNumId", namespaces=NS)

            if abstract_ref is not None:
                old_abstract_ref = abstract_ref.get(f"{W_NS}val")
                abstract_ref.set(
                    f"{W_NS}val",
                    abstract_id_remap.get(old_abstract_ref, old_abstract_ref),
                )

            if old_id in num_by_id:
                new_id = str(next_num_number)
                next_num_number += 1
                new_node.set(f"{W_NS}numId", new_id)
                num_id_remap[old_id] = new_id
            else:
                num_id_remap[old_id] = old_id

            num_by_id[num_id_remap[old_id]] = new_node
            ordered_num_ids.append(num_id_remap[old_id])

        merged_root = etree.Element(template_root.tag, nsmap=template_root.nsmap)

        for node in other_nodes:
            merged_root.append(node)

        for abstract_id in ordered_abstract_ids:
            merged_root.append(abstract_by_id[abstract_id])

        for num_id in ordered_num_ids:
            merged_root.append(num_by_id[num_id])

        merged_xml = etree.tostring(
            merged_root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

        return merged_xml, num_id_remap

    def _remap_numbering_ids(self, element, num_id_remap):

        if not num_id_remap:
            return

        for num_id_node in element.iter(f"{W_NS}numId"):
            value = num_id_node.get(f"{W_NS}val")
            mapped = num_id_remap.get(value)

            if mapped and mapped != value:
                num_id_node.set(f"{W_NS}val", mapped)

    # ------------------------------------------------------------------
    # Force Word to recalculate fields (TOC/STYLEREF/PAGEREF/SEQ) on open
    # ------------------------------------------------------------------

    SETTINGS_UPDATE_FIELDS_ANCHOR = "hdrShapeDefaults"

    def _build_settings_with_update_fields(self, base_docx_path):

        settings_part = "word/settings.xml"

        with ZipFile(base_docx_path, "r") as archive:
            if settings_part not in archive.namelist():
                return None
            settings_xml = archive.read(settings_part)

        root = etree.fromstring(settings_xml)

        if root.find("w:updateFields", namespaces=NS) is not None:
            return settings_xml

        update_fields = etree.Element(f"{W_NS}updateFields")
        update_fields.set(f"{W_NS}val", "true")

        anchor = root.find(f"w:{self.SETTINGS_UPDATE_FIELDS_ANCHOR}", namespaces=NS)

        if anchor is not None:
            anchor.addprevious(update_fields)
        else:
            root.insert(0, update_fields)

        return etree.tostring(
            root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

    def _build_custom_properties_xml(self, base_docx_path, resolved_metadata):

        custom_part = "docProps/custom.xml"

        with ZipFile(base_docx_path, "r") as archive:
            if custom_part not in archive.namelist():
                return None
            custom_xml = archive.read(custom_part)

        root = etree.fromstring(custom_xml)

        custom_ns = {
            "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
        }

        lowered_metadata = {
            key.lower(): value
            for key, value in resolved_metadata.items()
            if value is not None
        }

        changed = False

        for prop in root:
            name = prop.get("name")

            if not name:
                continue

            value = lowered_metadata.get(name.lower())

            if value is None:
                continue

            lpwstr = prop.find("vt:lpwstr", namespaces=custom_ns)

            if lpwstr is not None:
                lpwstr.text = value
                changed = True

        if not changed:
            return custom_xml

        return etree.tostring(
            root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

    def _build_core_properties_xml(self, source_docx_path, resolved_title):

        core_part = "docProps/core.xml"

        with ZipFile(source_docx_path, "r") as archive:
            if core_part not in archive.namelist():
                return None
            core_xml = archive.read(core_part)

        if not resolved_title:
            return core_xml

        root = etree.fromstring(core_xml)

        dc_ns = {"dc": "http://purl.org/dc/elements/1.1/"}
        title_el = root.find("dc:title", namespaces=dc_ns)

        if title_el is None:
            return core_xml

        title_el.text = resolved_title

        return etree.tostring(
            root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

    # ------------------------------------------------------------------
    # Cover metadata / field resolution
    # ------------------------------------------------------------------

    def _resolve_cover_metadata(self, model, overrides):

        config_path = self.project_root / "config" / "cover_metadata.json"
        defaults = {
            "Doc_Type": "Datasheet",
            "DocumentVersion": "Revision 1.0",
            "Title_continued": "",
            "Product_SalesCode": "",
            "Doc_Number": "",
        }

        if config_path.exists():
            with open(config_path, "r", encoding="utf-8") as handle:
                defaults.update(json.load(handle))

        metadata = dict(defaults)

        metadata["Date_of_document"] = metadata.get("Date_of_document") or date.today().isoformat()

        title = None

        if overrides:
            title = overrides.get("title") or overrides.get("Title")

        if not title:
            title = self._find_first_heading_text(model)

        if not title:
            title = Path(model.source_docx_path).stem

        metadata["Title"] = title

        if overrides:
            for key, value in overrides.items():
                if key.lower() == "title":
                    continue
                metadata[key] = value

        return metadata

    def _find_first_heading_text(self, model):

        ordered_items = sorted(model.content, key=lambda item: item.order_index)

        for item in ordered_items:
            if item.__class__.__name__.lower() == "heading" and getattr(item, "level", None) == 1:
                text = (getattr(item, "text", "") or "").strip()
                if text:
                    return text

        for item in ordered_items:
            if item.__class__.__name__.lower() == "heading":
                text = (getattr(item, "text", "") or "").strip()
                if text:
                    return text

        return None

    # ------------------------------------------------------------------
    # Cover-page placeholder content population
    # ------------------------------------------------------------------

    COVER_FEATURES_KEYWORDS = ("feature",)
    COVER_APPLICATIONS_KEYWORDS = ("application",)
    COVER_MAX_BULLETS = 10

    def _extract_cover_content(self, model):

        ordered_items = sorted(model.content, key=lambda item: item.order_index)

        return {
            "features": self._collect_bullets_after_heading(
                ordered_items, self.COVER_FEATURES_KEYWORDS
            )[: self.COVER_MAX_BULLETS],
            "applications": self._collect_bullets_after_heading(
                ordered_items, self.COVER_APPLICATIONS_KEYWORDS
            )[: self.COVER_MAX_BULLETS],
            "description": self._collect_description(ordered_items),
        }

    def _collect_bullets_after_heading(self, ordered_items, keywords):

        bullets = []
        collecting = False

        for item in ordered_items:
            class_name = item.__class__.__name__.lower()

            if class_name == "heading":
                if collecting:
                    break

                heading_text = (getattr(item, "text", "") or "").strip().lower()

                if any(keyword in heading_text for keyword in keywords):
                    collecting = True

                continue

            if collecting and class_name == "paragraph":
                text = (getattr(item, "text", "") or "").strip()
                if text:
                    bullets.append(text)

        return bullets

    def _collect_description(self, ordered_items):

        for idx, item in enumerate(ordered_items):

            if item.__class__.__name__.lower() != "heading":
                continue

            heading_text = (getattr(item, "text", "") or "").strip().lower()

            if "description" not in heading_text:
                continue

            for later_item in ordered_items[idx + 1:]:
                class_name = later_item.__class__.__name__.lower()

                if class_name == "heading":
                    break

                if class_name == "paragraph":
                    text = (getattr(later_item, "text", "") or "").strip()
                    if text:
                        return text

            break

        return None

    # ------------------------------------------------------------------
    # Cover-page block reordering (Description must precede Features/Applications)
    # ------------------------------------------------------------------

    COVER_BLOCK_ORDER = ["description", "features", "applications", "product_validation"]

    COVER_BLOCK_KEYWORDS = {
        "description": ("description",),
        "features": ("feature",),
        "applications": ("application",),
        "product_validation": ("product validation",),
    }

    def _reorder_cover_sections(self, elements):

        blocks = {}
        other_before = []
        other_after = []
        current_block = None
        seen_any_block = False
        collecting_ended = False

        for element in elements:

            if collecting_ended:
                other_after.append(element)
                continue

            is_heading = self._is_cover_heading_paragraph(element)
            block_key = self._match_cover_heading_block(element) if is_heading else None

            if is_heading and block_key is None:
                # A heading outside our known cover blocks (e.g. "Table of contents") ends reordering.
                collecting_ended = True
                other_after.append(element)
                continue

            if block_key is not None:
                current_block = block_key
                blocks.setdefault(current_block, []).append(element)
                seen_any_block = True
                continue

            if current_block is not None and etree.QName(element).localname == "p":
                blocks[current_block].append(element)
                continue

            if seen_any_block:
                other_after.append(element)
            else:
                other_before.append(element)

        if not blocks:
            return elements

        reordered = list(other_before)

        for key in self.COVER_BLOCK_ORDER:
            reordered.extend(blocks.get(key, []))

        reordered.extend(other_after)

        return reordered

    def _is_cover_heading_paragraph(self, element):

        if etree.QName(element).localname != "p":
            return False

        pstyle = element.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)

        return bool(pstyle) and pstyle[0] == "HeadingPreface"

    def _match_cover_heading_block(self, element):

        if not self._is_cover_heading_paragraph(element):
            return None

        text = "".join(element.xpath(".//*[local-name()='t']/text()")).strip().lower()

        for key, keywords in self.COVER_BLOCK_KEYWORDS.items():
            if any(keyword in text for keyword in keywords):
                return key

        return None

    def _build_cover_replacement_elements(self, element, cover_content):

        if etree.QName(element).localname != "p":
            return None

        text = "".join(element.xpath(".//*[local-name()='t']/text()")).strip().lower()

        if not text:
            return None

        if "insert a bulleted list of features" in text:
            return self._clone_bullet_paragraphs(element, cover_content.get("features"))

        if "insert a bulleted list of potential" in text:
            return self._clone_bullet_paragraphs(element, cover_content.get("applications"))

        if "description of the product" in text:
            description = cover_content.get("description")
            if not description:
                return None
            return [self._clone_paragraph_with_text(element, description)]

        return None

    def _clone_bullet_paragraphs(self, template_paragraph, texts):

        if not texts:
            return None

        return [
            self._clone_paragraph_with_text(template_paragraph, text)
            for text in texts
        ]

    def _clone_paragraph_with_text(self, template_paragraph, text):

        clone = deepcopy(template_paragraph)

        runs = clone.findall(f"{W_NS}r")
        template_rpr = runs[0].find(f"{W_NS}rPr") if runs else None

        for run in runs:
            clone.remove(run)

        for hyperlink in clone.findall(f"{W_NS}hyperlink"):
            clone.remove(hyperlink)

        clone.append(self._make_text_run(template_rpr, text))

        return clone

    def _make_text_run(self, template_rpr, text):

        run = etree.Element(f"{W_NS}r")

        if template_rpr is not None:
            run.append(deepcopy(template_rpr))

        t_node = etree.SubElement(run, f"{W_NS}t")
        t_node.text = text
        t_node.set(f"{XML_NS}space", "preserve")

        return run

    # ------------------------------------------------------------------
    # Template page geometry
    # ------------------------------------------------------------------

    def _get_template_content_width_twips(self, base_docx_path):

        with ZipFile(base_docx_path, "r") as archive:
            root = etree.fromstring(archive.read("word/document.xml"))

        sectpr_nodes = root.xpath("./w:body/w:sectPr", namespaces=NS)

        if not sectpr_nodes:
            sectpr_nodes = root.xpath(".//w:sectPr", namespaces=NS)

        if not sectpr_nodes:
            return None

        sectpr = sectpr_nodes[-1]

        pgsz = sectpr.find("w:pgSz", namespaces=NS)
        pgmar = sectpr.find("w:pgMar", namespaces=NS)

        if pgsz is None or pgmar is None:
            return None

        try:
            page_width = int(pgsz.get(f"{W_NS}w"))
            left_margin = int(pgmar.get(f"{W_NS}left", "0"))
            right_margin = int(pgmar.get(f"{W_NS}right", "0"))
        except (TypeError, ValueError):
            return None

        return max(page_width - left_margin - right_margin, 0)

    def _make_field_resolver(self, metadata, first_heading1_text):

        lowered_metadata = {key.lower(): value for key, value in metadata.items()}

        def resolve(instr_text):

            instr_norm = " ".join(instr_text.split())

            doc_property_match = re.search(
                r'DOCPROPERTY\s+"?([A-Za-z0-9_]+)"?',
                instr_norm,
                re.IGNORECASE,
            )

            if doc_property_match:
                key = doc_property_match.group(1).lower()
                return lowered_metadata.get(key)

            styleref_match = re.search(
                r'STYLEREF\s+"?([^"\\]+)"?',
                instr_norm,
                re.IGNORECASE,
            )

            if styleref_match:
                style_tokens = re.split(r"[;,]", styleref_match.group(1))
                normalized_tokens = [t.strip().lower().replace(" ", "") for t in style_tokens]

                if any(token == "heading1" for token in normalized_tokens):
                    return first_heading1_text

                return None

            return None

        return resolve

    def _update_fields_in_xml(self, xml_bytes, field_resolver):

        try:
            root = etree.fromstring(xml_bytes)
        except etree.XMLSyntaxError:
            return xml_bytes

        changed = False

        for paragraph in root.iter(f"{W_NS}p"):
            if self._update_fields_in_paragraph(paragraph, field_resolver):
                changed = True

        if not changed:
            return xml_bytes

        return etree.tostring(
            root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

    def _update_fields_in_paragraph(self, paragraph, field_resolver):

        changed = False

        for fld_simple in paragraph.findall(f".//{W_NS}fldSimple"):
            instr = fld_simple.get(f"{W_NS}instr", "")
            new_value = field_resolver(instr)

            if new_value is None:
                continue

            self._set_field_result_text(list(fld_simple), new_value)
            changed = True

        state = "idle"
        instr_parts = []
        result_runs = []

        for run in list(paragraph.findall(f"{W_NS}r")):

            fld_char = run.find(f"{W_NS}fldChar")

            if fld_char is not None:
                fld_type = fld_char.get(f"{W_NS}fldCharType")

                if fld_type == "begin":
                    state = "instr"
                    instr_parts = []
                    result_runs = []
                elif fld_type == "separate":
                    state = "result"
                elif fld_type == "end":
                    if state in {"instr", "result"} and instr_parts:
                        instr_text = "".join(instr_parts)
                        new_value = field_resolver(instr_text)

                        if new_value is not None and result_runs:
                            self._set_field_result_text(result_runs, new_value)
                            changed = True

                    state = "idle"
                    instr_parts = []
                    result_runs = []
                continue

            instr_node = run.find(f"{W_NS}instrText")

            if instr_node is not None and state == "instr":
                instr_parts.append(instr_node.text or "")
                continue

            if state == "result":
                result_runs.append(run)

        return changed

    def _set_field_result_text(self, runs, new_value):

        if not runs:
            return

        first_run = runs[0]
        t_node = first_run.find(f"{W_NS}t")

        if t_node is None:
            t_node = etree.SubElement(first_run, f"{W_NS}t")

        t_node.text = new_value
        t_node.set(f"{XML_NS}space", "preserve")

        for extra_run in runs[1:]:
            for extra_t in extra_run.findall(f"{W_NS}t"):
                extra_t.text = ""

    def _is_header_or_footer_part(self, filename):

        if filename.startswith("word/header") and filename.endswith(".xml"):
            return True

        if filename.startswith("word/footer") and filename.endswith(".xml"):
            return True

        return False

    # ------------------------------------------------------------------
    # Document rebuild: multi-section template replication
    # ------------------------------------------------------------------

    def _rebuild_document_xml(
        self,
        model,
        base_docx_path,
        preserve_template_cover,
        style_mapper,
        style_alias_map,
        id_remap,
        num_id_remap=None,
        cover_content=None,
        content_width_twips=None,
    ):

        with ZipFile(base_docx_path, "r") as source_zip:
            root = etree.fromstring(source_zip.read("word/document.xml"))

        body = root.xpath("./w:body", namespaces=NS)[0]

        sections = self._split_template_sections(body) if preserve_template_cover else None
        content_idx = self._find_content_section_index(sections) if sections else None

        for child in list(body):
            body.remove(child)

        ordered_items = sorted(model.content, key=lambda item: item.order_index)

        def append_migrated_content():
            for item in ordered_items:
                for element in self._elements_for_item(
                    item,
                    style_mapper=style_mapper,
                    style_alias_map=style_alias_map,
                    content_width_twips=content_width_twips,
                ):
                    self._remap_relationship_ids(element, id_remap)
                    self._remap_numbering_ids(element, num_id_remap or {})
                    body.append(element)

        if sections is not None and content_idx is not None:

            for idx in range(content_idx):

                section_elements = sections[idx][0]

                if idx == 0:
                    section_elements = self._reorder_cover_sections(section_elements)

                for element in section_elements:

                    replacement = self._build_cover_replacement_elements(
                        element,
                        cover_content or {},
                    )

                    if replacement is not None:
                        for repl_element in replacement:
                            body.append(repl_element)
                        continue

                    if self._is_forbidden_element(element):
                        continue

                    body.append(deepcopy(element))

            content_sectpr_template = sections[content_idx][1]

            append_migrated_content()

            has_suffix = content_idx < len(sections) - 1

            if has_suffix:

                if content_sectpr_template is not None:
                    body.append(self._make_section_break_paragraph(content_sectpr_template))

                for idx in range(content_idx + 1, len(sections)):
                    for element in sections[idx][0]:
                        if self._is_forbidden_element(element):
                            continue
                        body.append(deepcopy(element))

                final_sectpr = sections[-1][1]

                if final_sectpr is not None:
                    body.append(deepcopy(final_sectpr))

            else:
                if content_sectpr_template is not None:
                    body.append(deepcopy(content_sectpr_template))

        else:

            append_migrated_content()

            if model.body_sectpr_xml:
                body.append(etree.fromstring(model.body_sectpr_xml.encode("utf-8")))

        return etree.tostring(
            root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

    def _split_template_sections(self, body):
        """Split body children into an ordered list of (elements, sectPr_or_None)."""

        sections = []
        current = []

        for element in body:

            if etree.QName(element).localname == "sectPr":
                continue

            current.append(element)

            if etree.QName(element).localname == "p":
                sectpr_nodes = element.xpath("./w:pPr/w:sectPr", namespaces=NS)
                if sectpr_nodes:
                    sections.append((current, sectpr_nodes[0]))
                    current = []

        final_sectpr_nodes = body.xpath("./w:sectPr", namespaces=NS)
        final_sectpr = final_sectpr_nodes[0] if final_sectpr_nodes else None
        sections.append((current, final_sectpr))

        return sections

    def _find_content_section_index(self, sections):

        markers = self.formatting_policy.get(
            "content_section_markers",
            ["Formatting example", "{{MIGRATED_CONTENT}}"],
        )
        normalized_markers = {m.strip().lower() for m in markers}

        for idx, (elements, _sectpr) in enumerate(sections):
            for element in elements:
                if etree.QName(element).localname != "p":
                    continue

                text = "".join(element.xpath(".//*[local-name()='t']/text()")).strip()

                if text.lower() in normalized_markers:
                    return idx

        # Fallback: first section containing clearly instructional/example content.
        for idx, (elements, _sectpr) in enumerate(sections):
            for element in elements:
                if self._is_template_instructional_or_example_block(element):
                    return idx

        return None

    def _make_section_break_paragraph(self, sectpr):

        paragraph = etree.Element(f"{W_NS}p")
        ppr = etree.SubElement(paragraph, f"{W_NS}pPr")
        ppr.append(deepcopy(sectpr))
        return paragraph

    # ------------------------------------------------------------------
    # Content-item to XML element conversion
    # ------------------------------------------------------------------

    def _elements_for_item(
        self,
        item,
        style_mapper,
        style_alias_map,
        content_width_twips=None,
    ):

        if isinstance(item, FigureBlock):

            elements = []

            if item.caption_position == "before":
                elements.extend(
                    self._elements_for_xml(
                        item.caption.xml,
                        style_name=self._mapped_style_name(item.caption, style_mapper, item.caption.xml),
                        style_alias_map=style_alias_map,
                        apply_caption_numbering=True,
                    )
                )

            for image in item.images:
                elements.extend(
                    self._elements_for_xml(
                        image.xml,
                        style_alias_map=style_alias_map,
                    )
                )

            if item.caption_position == "after":
                elements.extend(
                    self._elements_for_xml(
                        item.caption.xml,
                        style_name=self._mapped_style_name(item.caption, style_mapper, item.caption.xml),
                        style_alias_map=style_alias_map,
                        apply_caption_numbering=True,
                    )
                )

            return elements

        if item.__class__.__name__.lower() == "table":
            return self._build_table_elements(
                item,
                style_mapper,
                style_alias_map,
                content_width_twips=content_width_twips,
            )

        is_caption = item.__class__.__name__.lower() == "caption"

        return self._elements_for_xml(
            item.xml,
            style_name=self._mapped_style_name(item, style_mapper, item.xml),
            style_alias_map=style_alias_map,
            apply_caption_numbering=is_caption,
        )

    def _elements_for_xml(
        self,
        xml_text,
        style_name=None,
        is_table=False,
        style_alias_map=None,
        apply_caption_numbering=False,
    ):

        element = etree.fromstring(xml_text.encode("utf-8"))

        # Section/page structure is owned exclusively by the template - never by the source.
        self._strip_embedded_section_break(element)

        if self._is_forbidden_element(element):
            return []

        style_id = self._resolve_style_id(style_name, style_alias_map)

        if style_id:
            if is_table:
                self._set_table_style(element, style_id)
            else:
                self._set_paragraph_style(element, style_id)

        if etree.QName(element).localname == "p":
            self._normalize_paragraph_formatting(element, style_id is not None)

            if style_id in HEADING_STYLE_IDS:
                # Chapter/section numbering is owned exclusively by the template's
                # heading styles - a direct numPr carried over from the source
                # (often numId="0", used by source authors to suppress their own
                # list numbering) would silently override and kill the template's
                # multilevel heading numbering, so it must never survive migration.
                self._strip_direct_paragraph_numbering(element)

            if apply_caption_numbering:
                self._apply_caption_number_bold(element)

        return [deepcopy(element)]

    def _strip_embedded_section_break(self, element):

        if etree.QName(element).localname != "p":
            return

        ppr = element.find("w:pPr", namespaces=NS)

        if ppr is None:
            return

        sectpr = ppr.find("w:sectPr", namespaces=NS)

        if sectpr is not None:
            ppr.remove(sectpr)

    def _strip_direct_paragraph_numbering(self, element):

        if etree.QName(element).localname != "p":
            return

        ppr = element.find("w:pPr", namespaces=NS)

        if ppr is None:
            return

        numpr = ppr.find("w:numPr", namespaces=NS)

        if numpr is not None:
            ppr.remove(numpr)

    def _mapped_style_name(
        self,
        item,
        style_mapper,
        xml_text,
    ):

        if style_mapper is None:
            return None

        item_name = item.__class__.__name__.lower()

        if item_name == "paragraph":
            if self._paragraph_has_numbering(xml_text):
                return style_mapper.style_for_list(
                    self._infer_list_kind(xml_text)
                )

        return style_mapper.style_for_item(item)

    # ------------------------------------------------------------------
    # Table styling
    # ------------------------------------------------------------------

    def _build_table_elements(self, item, style_mapper, style_alias_map, content_width_twips=None):

        element = etree.fromstring(item.xml.encode("utf-8"))

        if self._is_forbidden_element(element):
            return []

        table_style_id = self._resolve_real_table_style_id(style_alias_map)

        if table_style_id:
            self._set_table_style(element, table_style_id)

        self._ensure_table_look_banding(element)

        # Source colors/borders/widths must never survive; the template style drives appearance.
        self._strip_table_direct_formatting(element)
        self._fit_table_width_to_page(element, content_width_twips)

        if style_mapper is not None:
            self._apply_table_cell_styles(element, style_mapper, style_alias_map)

        return [deepcopy(element)]

    def _strip_table_direct_formatting(self, table_element):

        tblpr = table_element.find("w:tblPr", namespaces=NS)

        if tblpr is not None:
            for tag in ("w:tblBorders", "w:shd", "w:tblInd", "w:tblCellMar"):
                node = tblpr.find(tag, namespaces=NS)
                if node is not None:
                    tblpr.remove(node)

        for row in table_element.findall("w:tr", namespaces=NS):
            for cell in row.findall("w:tc", namespaces=NS):
                tcpr = cell.find("w:tcPr", namespaces=NS)

                if tcpr is None:
                    continue

                for tag in ("w:shd", "w:tcBorders", "w:tcMar"):
                    node = tcpr.find(tag, namespaces=NS)
                    if node is not None:
                        tcpr.remove(node)

    def _fit_table_width_to_page(self, table_element, content_width_twips):

        if not content_width_twips:
            return

        tblpr = table_element.find("w:tblPr", namespaces=NS)

        if tblpr is None:
            tblpr = etree.Element(f"{W_NS}tblPr")
            table_element.insert(0, tblpr)

        tblw = tblpr.find("w:tblW", namespaces=NS)

        if tblw is None:
            tblw = etree.SubElement(tblpr, f"{W_NS}tblW")

        tblw.set(f"{W_NS}w", str(content_width_twips))
        tblw.set(f"{W_NS}type", "dxa")

        tbllayout = tblpr.find("w:tblLayout", namespaces=NS)

        if tbllayout is None:
            tbllayout = etree.SubElement(tblpr, f"{W_NS}tblLayout")

        tbllayout.set(f"{W_NS}type", "fixed")

        grid = table_element.find("w:tblGrid", namespaces=NS)
        columns = grid.findall("w:gridCol", namespaces=NS) if grid is not None else []

        if not columns:
            return

        original_widths = []

        for col in columns:
            try:
                original_widths.append(int(col.get(f"{W_NS}w", "0")))
            except ValueError:
                original_widths.append(0)

        total_original = sum(original_widths) or 1

        new_widths = [
            max(round(width * content_width_twips / total_original), 1)
            for width in original_widths
        ]

        # Correct rounding drift so columns sum exactly to the available width.
        new_widths[-1] += content_width_twips - sum(new_widths)

        for col, width in zip(columns, new_widths):
            col.set(f"{W_NS}w", str(width))

        for row in table_element.findall("w:tr", namespaces=NS):

            col_idx = 0

            for cell in row.findall("w:tc", namespaces=NS):

                tcpr = cell.find("w:tcPr", namespaces=NS)
                span = 1

                if tcpr is not None:
                    gridspan = tcpr.find("w:gridSpan", namespaces=NS)
                    if gridspan is not None:
                        try:
                            span = max(int(gridspan.get(f"{W_NS}val", "1")), 1)
                        except ValueError:
                            span = 1

                cell_width = sum(new_widths[col_idx:col_idx + span])
                col_idx += span

                if tcpr is None:
                    tcpr = etree.Element(f"{W_NS}tcPr")
                    cell.insert(0, tcpr)

                tcw = tcpr.find("w:tcW", namespaces=NS)

                if tcw is None:
                    tcw = etree.SubElement(tcpr, f"{W_NS}tcW")

                tcw.set(f"{W_NS}w", str(cell_width))
                tcw.set(f"{W_NS}type", "dxa")

    def _resolve_real_table_style_id(self, style_alias_map):

        if not style_alias_map:
            return None

        for candidate in CANDIDATE_TABLE_STYLE_KEYS:
            resolved = style_alias_map.get(candidate)
            if resolved:
                return resolved

        return None

    def _ensure_table_look_banding(self, table_element):

        tblpr = table_element.find("w:tblPr", namespaces=NS)

        if tblpr is None:
            tblpr = etree.Element(f"{W_NS}tblPr")
            table_element.insert(0, tblpr)

        look = tblpr.find("w:tblLook", namespaces=NS)

        if look is None:
            look = etree.SubElement(tblpr, f"{W_NS}tblLook")

        look.set(f"{W_NS}firstRow", "1")
        look.set(f"{W_NS}lastRow", "0")
        look.set(f"{W_NS}firstColumn", "0")
        look.set(f"{W_NS}lastColumn", "0")
        look.set(f"{W_NS}noHBand", "0")
        look.set(f"{W_NS}noVBand", "1")
        look.set(f"{W_NS}val", "04A0")

    def _apply_table_cell_styles(self, table_element, style_mapper, style_alias_map):

        rows = table_element.findall("w:tr", namespaces=NS)

        for row_idx, row in enumerate(rows):

            is_header = row_idx == 0

            for cell in row.findall("w:tc", namespaces=NS):
                for paragraph in cell.findall(".//w:p", namespaces=NS):

                    alignment = self._paragraph_alignment(paragraph)
                    bold, italic = self._paragraph_has_bold_italic_run(paragraph)

                    style_key = style_mapper.style_for_table_cell(
                        is_header=is_header,
                        alignment=alignment,
                        bold=bold,
                        italic=italic,
                    )

                    if not style_key:
                        continue

                    resolved = self._resolve_style_id(style_key, style_alias_map)

                    if resolved:
                        self._set_paragraph_style(paragraph, resolved)

                    # Header/body cell text must use only the template style, never the source font/color.
                    self._normalize_paragraph_formatting(paragraph, True)

    def _paragraph_alignment(self, paragraph):

        jc = paragraph.xpath("./w:pPr/w:jc/@w:val", namespaces=NS)

        if jc and jc[0] in {"center", "both"}:
            return "center"

        return "left"

    def _paragraph_has_bold_italic_run(self, paragraph):

        bold = False
        italic = False

        for run in paragraph.findall(".//w:r", namespaces=NS):
            rpr = run.find("w:rPr", namespaces=NS)

            if rpr is None:
                continue

            if rpr.find("w:b", namespaces=NS) is not None:
                bold = True

            if rpr.find("w:i", namespaces=NS) is not None:
                italic = True

        return bold, italic

    # ------------------------------------------------------------------
    # Caption numbering
    # ------------------------------------------------------------------

    def _apply_caption_number_bold(self, paragraph_element):

        if self._paragraph_has_seq_field(paragraph_element):
            # A real Word Caption/SEQ field already exists (source authored via
            # Insert Caption) - preserve it exactly, only bold the label+field
            # portion, never flatten it to static text.
            self._bold_existing_caption_field(paragraph_element)
            return

        runs = paragraph_element.findall("w:r", namespaces=NS)

        if not runs:
            return

        combined = "".join(
            "".join(run.xpath(".//*[local-name()='t']/text()"))
            for run in runs
        )

        match = CAPTION_NUMBER_RE.match(combined)

        if not match:
            return

        prefix = match.group(0)
        rest = combined[len(prefix):]

        template_rpr = runs[0].find("w:rPr", namespaces=NS)
        label_match = re.match(r"(figure|table)", prefix, re.IGNORECASE)
        number_match = re.search(r"\d+(\.\d+)*", prefix)

        for run in runs:
            paragraph_element.remove(run)

        if label_match and number_match:
            # No pre-existing field - synthesize a real Word SEQ field so the
            # caption number updates automatically instead of being static text.
            label = label_match.group(1).capitalize()

            for new_run in self._build_caption_seq_field_runs(
                template_rpr,
                label=label,
                cached_number=number_match.group(0),
                remainder_text=rest,
            ):
                paragraph_element.append(new_run)
        else:
            paragraph_element.append(self._make_caption_run(template_rpr, prefix, bold=True))

            if rest:
                paragraph_element.append(self._make_caption_run(template_rpr, rest, bold=False))

    def _paragraph_has_seq_field(self, paragraph_element):

        instr_texts = paragraph_element.xpath(".//w:instrText/text()", namespaces=NS)
        return any("SEQ" in instr.upper() for instr in instr_texts)

    def _bold_existing_caption_field(self, paragraph_element):

        runs = paragraph_element.findall("w:r", namespaces=NS)
        end_index = None

        for idx, run in enumerate(runs):
            fldchar = run.find("w:fldChar", namespaces=NS)

            if fldchar is not None and fldchar.get(f"{W_NS}fldCharType") == "end":
                end_index = idx
                break

        if end_index is None:
            return

        for run in runs[: end_index + 1]:
            self._set_run_bold(run, True)

        for run in runs[end_index + 1 :]:
            self._set_run_bold(run, False)

    def _set_run_bold(self, run, bold):

        rpr = run.find("w:rPr", namespaces=NS)

        if rpr is None:
            rpr = etree.Element(f"{W_NS}rPr")
            run.insert(0, rpr)

        existing_bold = rpr.find("w:b", namespaces=NS)

        if bold:
            if existing_bold is None:
                etree.SubElement(rpr, f"{W_NS}b")
        elif existing_bold is not None:
            rpr.remove(existing_bold)

    def _build_caption_seq_field_runs(self, template_rpr, label, cached_number, remainder_text):

        def make_rpr(bold):
            rpr = deepcopy(template_rpr) if template_rpr is not None else etree.Element(f"{W_NS}rPr")
            existing_bold = rpr.find("w:b", namespaces=NS)

            if bold:
                if existing_bold is None:
                    etree.SubElement(rpr, f"{W_NS}b")
            elif existing_bold is not None:
                rpr.remove(existing_bold)

            return rpr

        def make_text_run(text, bold):
            run = etree.Element(f"{W_NS}r")
            run.append(make_rpr(bold))
            t_node = etree.SubElement(run, f"{W_NS}t")
            t_node.text = text
            t_node.set(f"{XML_NS}space", "preserve")
            return run

        def make_fldchar_run(fld_char_type, bold):
            run = etree.Element(f"{W_NS}r")
            run.append(make_rpr(bold))
            fldchar = etree.SubElement(run, f"{W_NS}fldChar")
            fldchar.set(f"{W_NS}fldCharType", fld_char_type)
            return run

        def make_instr_run(instr_text, bold):
            run = etree.Element(f"{W_NS}r")
            run.append(make_rpr(bold))
            instr_node = etree.SubElement(run, f"{W_NS}instrText")
            instr_node.text = instr_text
            instr_node.set(f"{XML_NS}space", "preserve")
            return run

        runs = [
            make_text_run(f"{label} ", bold=True),
            make_fldchar_run("begin", bold=True),
            make_instr_run(f" SEQ {label} \\* ARABIC ", bold=True),
            make_fldchar_run("separate", bold=True),
            make_text_run(cached_number, bold=True),
            make_fldchar_run("end", bold=True),
        ]

        if remainder_text:
            runs.append(make_text_run(remainder_text, bold=False))

        return runs

    def _make_caption_run(self, template_rpr, text, bold):

        run = etree.Element(f"{W_NS}r")
        rpr = deepcopy(template_rpr) if template_rpr is not None else etree.Element(f"{W_NS}rPr")

        existing_bold = rpr.find("w:b", namespaces=NS)

        if bold:
            if existing_bold is None:
                etree.SubElement(rpr, f"{W_NS}b")
        elif existing_bold is not None:
            rpr.remove(existing_bold)

        run.append(rpr)

        t_node = etree.SubElement(run, f"{W_NS}t")
        t_node.text = text
        t_node.set(f"{XML_NS}space", "preserve")

        return run

    # ------------------------------------------------------------------
    # Paragraph / table style application
    # ------------------------------------------------------------------

    def _set_paragraph_style(
        self,
        paragraph_element,
        style_name,
    ):

        if etree.QName(paragraph_element).localname != "p":
            return

        ppr = paragraph_element.find("w:pPr", namespaces=NS)

        if ppr is None:
            ppr = etree.Element(f"{W_NS}pPr")
            paragraph_element.insert(0, ppr)

        pstyle = ppr.find("w:pStyle", namespaces=NS)

        if pstyle is None:
            pstyle = etree.Element(f"{W_NS}pStyle")
            ppr.insert(0, pstyle)

        pstyle.set(f"{W_NS}val", style_name)

    def _set_table_style(
        self,
        table_element,
        style_name,
    ):

        if etree.QName(table_element).localname != "tbl":
            return

        tblpr = table_element.find("w:tblPr", namespaces=NS)

        if tblpr is None:
            tblpr = etree.Element(f"{W_NS}tblPr")
            table_element.insert(0, tblpr)

        tblstyle = tblpr.find("w:tblStyle", namespaces=NS)

        if tblstyle is None:
            tblstyle = etree.Element(f"{W_NS}tblStyle")
            tblpr.insert(0, tblstyle)

        tblstyle.set(f"{W_NS}val", style_name)

    # ------------------------------------------------------------------
    # Package part helpers
    # ------------------------------------------------------------------

    def _should_overlay_from_source(
        self,
        filename,
    ):

        if filename == "word/document.xml":
            return False

        template_owned_files = {
            "word/styles.xml",
            "word/stylesWithEffects.xml",
            "word/theme/theme1.xml",
            "word/numbering.xml",
            "word/settings.xml",
        }

        if filename in {"[Content_Types].xml", "word/_rels/document.xml.rels"}:
            return False

        if self._is_header_or_footer_part(filename):
            return False

        if filename.startswith("word/_rels/header") and filename.endswith(".rels"):
            return False

        if filename.startswith("word/_rels/footer") and filename.endswith(".rels"):
            return False

        return filename not in template_owned_files

    def _merge_content_types(self, template_types_xml, source_types_xml):

        template_root = etree.fromstring(template_types_xml)
        source_root = etree.fromstring(source_types_xml)

        merged_root = etree.Element(template_root.tag, nsmap=template_root.nsmap)

        defaults = {}
        overrides = {}

        for node in template_root:
            local_name = etree.QName(node).localname

            if local_name == "Default":
                key = node.get("Extension", "").lower()
                defaults[key] = deepcopy(node)
            elif local_name == "Override":
                key = node.get("PartName", "")
                overrides[key] = deepcopy(node)

        for node in source_root:
            local_name = etree.QName(node).localname

            if local_name == "Default":
                key = node.get("Extension", "").lower()
                defaults[key] = deepcopy(node)
            elif local_name == "Override":
                key = node.get("PartName", "")
                overrides[key] = deepcopy(node)

        fallback_defaults = {
            "png": "image/png",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "gif": "image/gif",
            "bmp": "image/bmp",
            "tif": "image/tiff",
            "tiff": "image/tiff",
            "emf": "image/x-emf",
            "wmf": "image/x-wmf",
        }

        for extension, content_type in fallback_defaults.items():
            if extension in defaults:
                continue

            default_node = etree.Element(f"{{{template_root.nsmap[None]}}}Default")
            default_node.set("Extension", extension)
            default_node.set("ContentType", content_type)
            defaults[extension] = default_node

        for key in sorted(defaults.keys()):
            merged_root.append(defaults[key])

        for key in sorted(overrides.keys()):
            merged_root.append(overrides[key])

        return etree.tostring(
            merged_root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

    def _load_formatting_policy(self):

        policy_path = self.project_root / "config" / "formatting_policy.json"

        if not policy_path.exists():
            return {}

        with open(policy_path, "r", encoding="utf-8") as handle:
            return json.load(handle)

    def _load_template_style_aliases(self, template_path):

        aliases = {}

        with ZipFile(template_path, "r") as archive:
            if "word/styles.xml" not in archive.namelist():
                return aliases

            styles_root = etree.fromstring(archive.read("word/styles.xml"))

        style_nodes = styles_root.xpath(".//*[local-name()='style']")

        for style in style_nodes:
            style_id = style.get(f"{W_NS}styleId")

            if not style_id:
                continue

            aliases[style_id.lower()] = style_id

            name_values = style.xpath("./*[local-name()='name']/@*[local-name()='val']")

            if name_values:
                aliases[name_values[0].strip().lower()] = style_id

        return aliases

    def _resolve_style_id(self, style_name, style_alias_map):

        if not style_name:
            return None

        if style_alias_map is None:
            return style_name

        resolved = style_alias_map.get(style_name.strip().lower())

        if resolved:
            return resolved

        return style_name

    # ------------------------------------------------------------------
    # Formatting normalization
    # ------------------------------------------------------------------

    def _normalize_paragraph_formatting(self, paragraph_element, is_mapped_block):

        if not is_mapped_block:
            return

        strip_runs = self.formatting_policy.get(
            "strip_direct_run_formatting_for_mapped_blocks",
            False,
        )

        normalize_spacing = self.formatting_policy.get(
            "normalize_paragraph_spacing_to_template",
            False,
        )

        if normalize_spacing:
            ppr = paragraph_element.find("w:pPr", namespaces=NS)

            if ppr is not None:
                for tag in ("w:spacing", "w:ind"):
                    node = ppr.find(tag, namespaces=NS)
                    if node is not None:
                        ppr.remove(node)

        if not strip_runs:
            return

        preserve_bold = self.formatting_policy.get("preserve_bold", True)
        preserve_italic = self.formatting_policy.get("preserve_italic", True)
        preserve_underline = self.formatting_policy.get("preserve_underline", False)

        for run in paragraph_element.xpath(".//*[local-name()='r']"):

            rpr = run.find("w:rPr", namespaces=NS)

            if rpr is None:
                continue

            has_bold = rpr.find("w:b", namespaces=NS) is not None
            has_italic = rpr.find("w:i", namespaces=NS) is not None
            underline_node = rpr.find("w:u", namespaces=NS)

            run.remove(rpr)

            new_rpr = etree.Element(f"{W_NS}rPr")

            if preserve_bold and has_bold:
                new_rpr.append(etree.Element(f"{W_NS}b"))

            if preserve_italic and has_italic:
                new_rpr.append(etree.Element(f"{W_NS}i"))

            if preserve_underline and underline_node is not None:
                new_underline = etree.Element(f"{W_NS}u")
                underline_val = underline_node.get(f"{W_NS}val")
                if underline_val:
                    new_underline.set(f"{W_NS}val", underline_val)
                new_rpr.append(new_underline)

            if len(new_rpr):
                run.insert(0, new_rpr)

    def _is_forbidden_element(self, element):

        local_name = etree.QName(element).localname

        if local_name == "tbl":
            table_text = " ".join(element.xpath(".//*[local-name()='t']/text()")).strip().lower()
            forbidden_list = self.formatting_policy.get("forbidden_content", [])
            return any(phrase.lower() in table_text for phrase in forbidden_list)

        if local_name != "p":
            return False

        forbidden_list = self.formatting_policy.get("forbidden_content", [])

        if not forbidden_list:
            return False

        text = "".join(element.xpath(".//*[local-name()='t']/text()")).strip().lower()

        if not text:
            return False

        return any(phrase.lower() in text for phrase in forbidden_list)

    def _paragraph_has_numbering(self, xml_text):

        element = etree.fromstring(xml_text.encode("utf-8"))

        if etree.QName(element).localname != "p":
            return False

        return bool(element.xpath("./*[local-name()='pPr']/*[local-name()='numPr']"))

    def _infer_list_kind(self, xml_text):

        element = etree.fromstring(xml_text.encode("utf-8"))

        style_vals = element.xpath(
            "./*[local-name()='pPr']/*[local-name()='pStyle']/@*[local-name()='val']"
        )

        if style_vals:
            style_hint = style_vals[0].lower()
            if "number" in style_hint:
                return "number"
            if "bullet" in style_hint:
                return "bullet"

        text = "".join(element.xpath(".//*[local-name()='t']/text()")).strip()

        if text and (text[0].isdigit() or text[:2].lower() in {"i.", "a."}):
            return "number"

        return "bullet"

    def _is_template_instructional_or_example_block(self, element):

        local_name = etree.QName(element).localname

        if local_name == "tbl":
            table_text = " ".join(element.xpath(".//*[local-name()='t']/text()")).strip().lower()
            return any(
                token in table_text
                for token in {"example", "placeholder", "sample", "instruction"}
            )

        if local_name != "p":
            return False

        text = "".join(element.xpath(".//*[local-name()='t']/text()")).strip().lower()

        if not text:
            return False

        if self._is_forbidden_element(element):
            return True

        instructional_tokens = {
            "insert",
            "placeholder",
            "example",
            "sample",
            "instruction",
            "lorem ipsum",
            "dummy",
            "to be replaced",
        }

        if any(token in text for token in instructional_tokens):
            return True

        return False
