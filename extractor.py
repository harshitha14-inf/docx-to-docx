import re
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.enum.section import WD_ORIENT
from lxml import etree

from models import (
    Caption,
    Cell,
    DocumentModel,
    FigureBlock,
    Footer,
    Formula,
    Heading,
    Header,
    Image,
    ImageRef,
    Paragraph,
    RawBlock,
    Run,
    SectionInfo,
    SectionMargins,
    Table,
    TextBox,
)


CAPTION_RE = re.compile(
    r"^(figure|fig\.)\s+\d+(?:\.\d+)*(?::|\b)",
    re.IGNORECASE,
)

TABLE_CAPTION_RE = re.compile(
    r"^table\s+\d+(?:\.\d+)*(?::|\b)",
    re.IGNORECASE,
)

# Formula detection must run before any heading/paragraph classification -
# equations are content, never headings, regardless of the source paragraph's
# style or outline level (see repo memory for the corruption this caused).
FORMULA_FUNCTION_RE = re.compile(
    r"\b(ATAN2|ATAN|SIN|COS|TAN|SQRT|LOG)\s*\(",
    re.IGNORECASE,
)

FORMULA_ASSIGNMENT_RE = re.compile(
    r"[A-Za-z_][A-Za-z0-9_]*(?:<[^>\s]{1,20}>)?\s*(?:\+=|-=|\*=|/=|=)\s*[\w(<\-.]"
)


class Extractor:

    def __init__(self, docx_path):
        self.docx_path = Path(docx_path)

    @property
    def _namespaces(self):
        return {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        }

    def extract(self):

        document = Document(
            self.docx_path
        )

        model = DocumentModel(
            source_docx_path=str(
                self.docx_path
            )
        )

        self._extract_headers_and_footers(
            document,
            model,
        )

        model.sections = self._extract_sections(
            document
        )

        media_dir = Path(
            "extracted_images"
        )
        media_dir.mkdir(
            exist_ok=True
        )

        with ZipFile(
            self.docx_path
        ) as archive:

            image_paths = self._extract_media_files(
                archive,
                media_dir,
            )

            rid_to_target = self._extract_relationships(
                archive
            )

            root = etree.fromstring(
                archive.read(
                    "word/document.xml"
                )
            )

        body = root.xpath(
            "./w:body",
            namespaces=self._namespaces,
        )[0]

        body_sectpr = body.xpath(
            "./w:sectPr",
            namespaces=self._namespaces,
        )

        if body_sectpr:
            model.body_sectpr_xml = etree.tostring(
                body_sectpr[0],
                encoding="unicode",
            )

        provisional_items = []
        order_index = 0

        for element in body:

            tag = etree.QName(
                element
            ).localname

            if tag == "sectPr":
                continue

            if tag == "p":
                provisional_items.append(
                    self._extract_paragraph_item(
                        element,
                        image_paths,
                        rid_to_target,
                        order_index,
                    )
                )
                order_index += 1
                continue

            if tag == "tbl":
                provisional_items.append(
                    self._extract_table_item(
                        element,
                        order_index,
                    )
                )
                order_index += 1
                continue

            provisional_items.append(
                RawBlock(
                    order_index=order_index,
                    xml=etree.tostring(
                        element,
                        encoding="unicode",
                    ),
                    block_type=tag,
                )
            )
            order_index += 1

        model.content = self._group_figure_blocks(
            provisional_items
        )

        self._mark_formula_grouping(
            model.content
        )

        return model

    def _extract_headers_and_footers(
        self,
        document,
        model,
    ):

        for section in document.sections:

            header_text = "\n".join(
                paragraph.text.strip()
                for paragraph in section.header.paragraphs
                if paragraph.text.strip()
            )

            if header_text:
                header_style = None

                for paragraph in section.header.paragraphs:
                    if paragraph.text.strip() and paragraph.style is not None:
                        header_style = paragraph.style.name
                        break

                model.headers.append(
                    Header(
                        text=header_text,
                        style=header_style,
                    )
                )

            footer_text = "\n".join(
                paragraph.text.strip()
                for paragraph in section.footer.paragraphs
                if paragraph.text.strip()
            )

            if footer_text:
                footer_style = None

                for paragraph in section.footer.paragraphs:
                    if paragraph.text.strip() and paragraph.style is not None:
                        footer_style = paragraph.style.name
                        break

                model.footers.append(
                    Footer(
                        text=footer_text,
                        style=footer_style,
                    )
                )

    def _extract_sections(
        self,
        document,
    ):

        sections = []

        for section in document.sections:

            orientation = "portrait"

            if section.orientation == WD_ORIENT.LANDSCAPE:
                orientation = "landscape"

            sections.append(
                SectionInfo(
                    orientation=orientation,
                    page_width=self._length_to_int(
                        section.page_width
                    ),
                    page_height=self._length_to_int(
                        section.page_height
                    ),
                    start_type=self._enum_name(
                        section.start_type
                    ),
                    margins=SectionMargins(
                        left=self._length_to_int(
                            section.left_margin
                        ),
                        right=self._length_to_int(
                            section.right_margin
                        ),
                        top=self._length_to_int(
                            section.top_margin
                        ),
                        bottom=self._length_to_int(
                            section.bottom_margin
                        ),
                        header=self._length_to_int(
                            section.header_distance
                        ),
                        footer=self._length_to_int(
                            section.footer_distance
                        ),
                        gutter=self._length_to_int(
                            section.gutter
                        ),
                    ),
                )
            )

        return sections

    def _extract_media_files(
        self,
        archive,
        media_dir,
    ):

        image_paths = {}

        for media_file in archive.namelist():

            if not media_file.startswith(
                "word/media/"
            ):
                continue

            image_name = Path(
                media_file
            ).name

            image_path = (
                media_dir /
                image_name
            )

            with archive.open(
                media_file
            ) as src:

                with open(
                    image_path,
                    "wb",
                ) as dst:

                    dst.write(
                        src.read()
                    )

            image_paths[
                image_name
            ] = str(image_path)

        return image_paths

    def _extract_relationships(
        self,
        archive,
    ):

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

    def _extract_paragraph_item(
        self,
        element,
        image_paths,
        rid_to_target,
        order_index,
    ):

        xml = etree.tostring(
            element,
            encoding="unicode",
        )

        image_refs = self._get_image_refs(
            element,
            image_paths,
            rid_to_target,
            order_index,
        )

        if image_refs:
            return Image(
                order_index=order_index,
                xml=xml,
                refs=image_refs,
            )

        textbox_texts = self._get_textbox_texts(
            element
        )

        text = self._get_visible_paragraph_text(
            element
        ).strip()

        if textbox_texts and not text:
            if any(self._is_formula_text(t) for t in textbox_texts):
                return Formula(
                    order_index=order_index,
                    xml=xml,
                    text="\n".join(textbox_texts),
                )
            return TextBox(
                order_index=order_index,
                xml=xml,
                texts=textbox_texts,
            )

        if CAPTION_RE.match(text):
            return Caption(
                order_index=order_index,
                xml=xml,
                text=text,
                caption_type="image",
                style=self._get_paragraph_style_name(element),
            )

        if TABLE_CAPTION_RE.match(text):
            return Caption(
                order_index=order_index,
                xml=xml,
                text=text,
                caption_type="table",
                style=self._get_paragraph_style_name(element),
            )

        # Formula detection runs before heading classification - a source
        # paragraph carrying a Heading style/outline level whose text is an
        # equation must still become Formula content, never a Heading.
        if self._is_formula_text(text):
            return Formula(
                order_index=order_index,
                xml=xml,
                text=text,
                style=self._get_paragraph_style_name(element),
                runs=self._extract_runs(element),
            )

        heading_level = self._detect_heading_level(
            element
        )

        runs = self._extract_runs(element)

        if heading_level is not None:
            return Heading(
                order_index=order_index,
                xml=xml,
                level=heading_level,
                text=text,
                style=self._get_paragraph_style_name(element),
                runs=runs,
            )

        return Paragraph(
            order_index=order_index,
            xml=xml,
            text=text,
            style=self._get_paragraph_style_name(element),
            runs=runs,
        )

    def _extract_table_item(
        self,
        element,
        order_index,
    ):

        table_data = []
        table_cells = []

        rows = element.xpath(
            ".//*[local-name()='tr']"
        )

        for row_idx, row in enumerate(rows):

            row_data = []
            row_cells = []

            for cell in row.xpath(
                ".//*[local-name()='tc']"
            ):

                value = " ".join(
                    cell.xpath(
                        ".//*[local-name()='t']/text()"
                    )
                )

                row_data.append(
                    value.strip()
                )

                row_cells.append(
                    Cell(
                        text=value.strip(),
                        runs=self._extract_runs(cell),
                        row_span=self._get_row_span(cell),
                        col_span=self._get_col_span(cell),
                        is_header=(row_idx == 0),
                        raw_xml=etree.tostring(cell, encoding="unicode"),
                    )
                )

            table_data.append(
                row_data
            )
            table_cells.append(
                row_cells
            )

        return Table(
            order_index=order_index,
            xml=etree.tostring(
                element,
                encoding="unicode",
            ),
            data=table_data,
            cells=table_cells,
        )

    def _get_col_span(
        self,
        cell,
    ):

        gridspan = cell.xpath(
            "./*[local-name()='tcPr']/*[local-name()='gridSpan']/@*[local-name()='val']"
        )

        if not gridspan:
            return 1

        try:
            return max(int(gridspan[0]), 1)
        except ValueError:
            return 1

    def _get_row_span(
        self,
        cell,
    ):

        vmerge = cell.xpath(
            "./*[local-name()='tcPr']/*[local-name()='vMerge']/@*[local-name()='val']"
        )

        if vmerge and vmerge[0] == "continue":
            return 0

        return 1

    def _extract_runs(
        self,
        element,
    ):

        runs = []

        for run in element.xpath(
            ".//*[local-name()='r']"
        ):

            ancestor = run.getparent()
            skip_run = False

            while ancestor is not None and ancestor is not element:

                local_name = etree.QName(
                    ancestor
                ).localname

                if local_name in {
                    "drawing",
                    "txbxContent",
                }:
                    skip_run = True
                    break

                ancestor = ancestor.getparent()

            if skip_run:
                continue

            text = "".join(
                run.xpath(
                    "./*[local-name()='t']/text()"
                )
            )

            if not text:
                continue

            rpr = run.find(
                f"{{{self._namespaces['w']}}}rPr"
            )

            bold = False
            italic = False
            underline = False

            if rpr is not None:
                bold = rpr.find(f"{{{self._namespaces['w']}}}b") is not None
                italic = rpr.find(f"{{{self._namespaces['w']}}}i") is not None
                underline_node = rpr.find(f"{{{self._namespaces['w']}}}u")
                underline = (
                    underline_node is not None
                    and underline_node.get(f"{{{self._namespaces['w']}}}val") not in (None, "none")
                )

            hyperlink_url = None

            hyperlink_ancestor = run.xpath(
                "ancestor::*[local-name()='hyperlink'][1]"
            )

            if hyperlink_ancestor:
                hyperlink_url = hyperlink_ancestor[0].get(
                    f"{{{self._namespaces['r']}}}id"
                )

            runs.append(
                Run(
                    text=text,
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    hyperlink_url=hyperlink_url,
                )
            )

        return runs

    def _get_image_refs(
        self,
        element,
        image_paths,
        rid_to_target,
        order_index,
    ):

        relationship_ids = []

        for r_id in element.xpath(
            ".//*[local-name()='blip']/@r:embed",
            namespaces=self._namespaces,
        ):
            relationship_ids.append(r_id)

        for r_id in element.xpath(
            ".//*[local-name()='imagedata']/@r:id",
            namespaces=self._namespaces,
        ):
            relationship_ids.append(r_id)

        seen = set()
        image_refs = []

        for r_id in relationship_ids:

            if r_id in seen:
                continue

            seen.add(r_id)
            target = rid_to_target.get(
                r_id,
                "",
            )
            image_name = Path(target).name

            image_refs.append(
                ImageRef(
                    name=image_name,
                    path=image_paths.get(
                        image_name,
                        "",
                    ),
                    relationship_id=r_id,
                    image_format=Path(target).suffix.lower().lstrip("."),
                    document_position_index=order_index,
                    target=target,
                )
            )

        return image_refs

    def _get_textbox_texts(
        self,
        element,
    ):

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

    def _get_visible_paragraph_text(
        self,
        element,
    ):

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

    def _group_figure_blocks(
        self,
        items,
    ):

        grouped_items = []
        index = 0

        while index < len(items):

            item = items[index]

            if (
                isinstance(item, Caption)
                and item.caption_type == "image"
            ):

                images = []
                look_ahead = index + 1

                while (
                    look_ahead < len(items)
                    and isinstance(
                        items[look_ahead],
                        Image,
                    )
                ):
                    images.append(
                        items[look_ahead]
                    )
                    look_ahead += 1

                if images:
                    grouped_items.append(
                        FigureBlock(
                            order_index=item.order_index,
                            xml="",
                            caption=item,
                            images=images,
                            caption_position="before",
                        )
                    )
                    index = look_ahead
                    continue

            if isinstance(item, Image):

                images = [item]
                look_ahead = index + 1

                while (
                    look_ahead < len(items)
                    and isinstance(
                        items[look_ahead],
                        Image,
                    )
                ):
                    images.append(
                        items[look_ahead]
                    )
                    look_ahead += 1

                if (
                    look_ahead < len(items)
                    and isinstance(
                        items[look_ahead],
                        Caption,
                    )
                    and items[look_ahead].caption_type == "image"
                ):
                    grouped_items.append(
                        FigureBlock(
                            order_index=item.order_index,
                            xml="",
                            caption=items[look_ahead],
                            images=images,
                            caption_position="after",
                        )
                    )
                    index = look_ahead + 1
                    continue

            grouped_items.append(item)
            index += 1

        return grouped_items

    def _is_formula_text(
        self,
        text,
    ):

        if not text:
            return False

        if FORMULA_FUNCTION_RE.search(text):
            return True

        # Anchored at the start - an equation line IS an assignment, whereas
        # a normal sentence that merely mentions a value inline ("operating
        # temperature = -40 to +150\u00b0C") must stay a plain paragraph.
        if FORMULA_ASSIGNMENT_RE.match(text.strip()):
            return True

        return False

    def _mark_formula_grouping(
        self,
        items,
    ):
        """Flag the item immediately preceding a Formula for keep-with-next.

        Preserves the explanation-paragraph -> formula -> formula-note visual
        group (page-break/paragraph-flow only - order is already preserved by
        order_index, this only prevents Word from splitting the group across
        a page boundary).
        """

        for index in range(len(items) - 1):

            if not isinstance(items[index + 1], Formula):
                continue

            current_item = items[index]

            if isinstance(current_item, (Paragraph, Heading, Formula)):
                current_item.keep_with_next = True

    def _enum_name(
        self,
        enum_value,
    ):

        name = getattr(
            enum_value,
            "name",
            None,
        )

        if name:
            return name.lower()

        return str(
            enum_value
        ).lower()

    def _length_to_int(
        self,
        value,
    ):

        if value is None:
            return 0

        return int(value)

    def _get_paragraph_style_name(
        self,
        element,
    ):

        style_value = element.xpath(
            "./*[local-name()='pPr']/*[local-name()='pStyle']/@*[local-name()='val']"
        )

        if not style_value:
            return None

        return style_value[0]

    def _detect_heading_level(
        self,
        element,
    ):

        style_name = (
            self._get_paragraph_style_name(
                element
            )
            or ""
        ).strip().lower()

        match = re.search(
            r"heading\s*([1-9])$",
            style_name,
        )

        if match:
            return int(match.group(1))

        outline_levels = element.xpath(
            "./*[local-name()='pPr']/*[local-name()='outlineLvl']/@*[local-name()='val']"
        )

        if outline_levels:
            try:
                # Word stores outline level as zero-based.
                return int(outline_levels[0]) + 1
            except ValueError:
                return None

        return None