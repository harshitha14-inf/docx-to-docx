import re
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.enum.section import WD_ORIENT
from lxml import etree

from models import (
    Caption,
    DocumentModel,
    FigureBlock,
    Footer,
    Heading,
    Header,
    Image,
    ImageRef,
    Paragraph,
    RawBlock,
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

        heading_level = self._detect_heading_level(
            element
        )

        if heading_level is not None:
            return Heading(
                order_index=order_index,
                xml=xml,
                level=heading_level,
                text=text,
                style=self._get_paragraph_style_name(element),
            )

        return Paragraph(
            order_index=order_index,
            xml=xml,
            text=text,
            style=self._get_paragraph_style_name(element),
        )

    def _extract_table_item(
        self,
        element,
        order_index,
    ):

        table_data = []

        rows = element.xpath(
            ".//*[local-name()='tr']"
        )

        for row in rows:

            row_data = []

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

            table_data.append(
                row_data
            )

        return Table(
            order_index=order_index,
            xml=etree.tostring(
                element,
                encoding="unicode",
            ),
            data=table_data,
        )

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