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

        if template_used:
            merged_rels_bytes, id_remap = self._compute_merged_relationships(
                base_docx_path=base_docx_path,
                source_docx_path=model.source_docx_path,
            )

        resolved_cover_metadata = self._resolve_cover_metadata(
            model,
            overrides=cover_metadata,
        )

        first_heading_text = self._find_first_heading_text(model)

        field_resolver = self._make_field_resolver(
            resolved_cover_metadata,
            first_heading_text,
        )

        rebuilt_document = self._rebuild_document_xml(
            model,
            base_docx_path=base_docx_path,
            preserve_template_cover=preserve_template_cover and template_used,
            style_mapper=active_style_mapper,
            style_alias_map=style_alias_map,
            id_remap=id_remap,
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
                ):
                    self._remap_relationship_ids(element, id_remap)
                    body.append(element)

        if sections is not None and content_idx is not None:

            for idx in range(content_idx):
                for element in sections[idx][0]:
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
            return self._build_table_elements(item, style_mapper, style_alias_map)

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

            if apply_caption_numbering:
                self._apply_caption_number_bold(element)

        return [deepcopy(element)]

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

    def _build_table_elements(self, item, style_mapper, style_alias_map):

        element = etree.fromstring(item.xml.encode("utf-8"))

        if self._is_forbidden_element(element):
            return []

        table_style_id = self._resolve_real_table_style_id(style_alias_map)

        if table_style_id:
            self._set_table_style(element, table_style_id)

        self._ensure_table_look_banding(element)

        if style_mapper is not None:
            self._apply_table_cell_styles(element, style_mapper, style_alias_map)

        return [deepcopy(element)]

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

        for run in runs:
            paragraph_element.remove(run)

        paragraph_element.append(self._make_caption_run(template_rpr, prefix, bold=True))

        if rest:
            paragraph_element.append(self._make_caption_run(template_rpr, rest, bold=False))

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
                spacing = ppr.find("w:spacing", namespaces=NS)
                if spacing is not None:
                    ppr.remove(spacing)

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
