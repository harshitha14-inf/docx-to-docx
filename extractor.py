import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

from docx import Document

from models import (
    DocumentModel,
    Paragraph,
    Table,
    Image,
    Caption,
    Header,
    Footer
)


class Extractor:

    def __init__(self, docx_path):
        self.docx_path = docx_path

    def extract(self):

        document = Document(
            self.docx_path
        )

        model = DocumentModel()

        # ------------------------
        # Header Extraction
        # ------------------------

        for section in document.sections:

            header_text = "\n".join(

                p.text.strip()

                for p in section.header.paragraphs

                if p.text.strip()
            )

            if header_text:

                model.headers.append(
                    Header(header_text)
                )

            footer_text = "\n".join(

                p.text.strip()

                for p in section.footer.paragraphs

                if p.text.strip()
            )

            if footer_text:

                model.footers.append(
                    Footer(footer_text)
                )

        # ------------------------
        # Images
        # ------------------------

        media_dir = Path(
            "extracted_images"
        )

        media_dir.mkdir(
            exist_ok=True
        )

        extracted_images = []
        image_by_name = {}

        # rId -> image filename, built from document relationships
        rid_to_image = {}

        with ZipFile(
            self.docx_path
        ) as z:

            media_files = [

                f

                for f in z.namelist()

                if f.startswith(
                    "word/media/"
                )
            ]

            for media_file in media_files:

                image_name = Path(
                    media_file
                ).name

                image_path = (
                    media_dir /
                    image_name
                )

                with z.open(
                    media_file
                ) as src:

                    with open(

                        image_path,

                        "wb"

                    ) as dst:

                        dst.write(
                            src.read()
                        )

                img = Image(
                    name=image_name,
                    path=str(image_path)
                )

                extracted_images.append(img)
                image_by_name[image_name] = img

            rels_path = "word/_rels/document.xml.rels"

            if rels_path in z.namelist():

                rels_root = ET.fromstring(
                    z.read(rels_path)
                )

                for rel in rels_root:

                    r_id = rel.get("Id")
                    target = rel.get("Target", "")

                    if target.startswith("media/"):
                        rid_to_image[r_id] = (
                            Path(target).name
                        )

        # ------------------------
        # Ordered Content
        # ------------------------

        body = document.element.body

        for element in body:

            tag = element.tag.split(
                "}"
            )[-1]

            # --------------------
            # Paragraph
            # --------------------

            if tag == "p":

                drawing_nodes = element.xpath(
                    ".//*[local-name()='drawing']"
                )

                if drawing_nodes:

                    embed_attrs = element.xpath(
                        ".//*[local-name()='blip']"
                        "/@*[local-name()='embed']"
                    )

                    # one Image per blip; skip drawings with no image (charts, shapes)
                    for r_id in embed_attrs:

                        img_name = rid_to_image.get(r_id)

                        if img_name:

                            img = image_by_name.get(img_name)

                            if img is not None:
                                model.content.append(img)

                W_NS = (
                    "http://schemas.openxmlformats.org"
                    "/wordprocessingml/2006/main"
                )
                W_T = f"{{{W_NS}}}t"
                W_DRAWING = f"{{{W_NS}}}drawing"

                text_parts = []

                for t_node in element.iter(W_T):

                    if not t_node.text:
                        continue

                    # skip w:t nodes that live inside a w:drawing
                    ancestor = t_node.getparent()
                    in_drawing = False

                    while ancestor is not None and ancestor is not element:
                        if ancestor.tag == W_DRAWING:
                            in_drawing = True
                            break
                        ancestor = ancestor.getparent()

                    if not in_drawing:
                        text_parts.append(t_node.text)

                text = "".join(
                    text_parts
                ).strip()

                if text:

                    lower = text.lower()

                    # Caption Detection
                    # ----------------

                    if (
                        lower.startswith(
                            "figure "
                        )
                        or lower.startswith(
                            "fig."
                        )
                    ):

                        model.content.append(

                            Caption(

                                text=text,

                                caption_type="image"
                            )
                        )

                    elif (
                        lower.startswith(
                            "table "
                        )
                    ):

                        model.content.append(

                            Caption(

                                text=text,

                                caption_type="table"
                            )
                        )

                    else:

                        model.content.append(
                            Paragraph(
                                text
                            )
                        )

            # --------------------
            # Table
            # --------------------

            elif tag == "tbl":

                table_data = []

                rows = element.xpath(
                    ".//*[local-name()='tr']"
                )

                for row in rows:

                    row_data = []

                    cells = row.xpath(
                        ".//*[local-name()='tc']"
                    )

                    for cell in cells:

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

                model.content.append(
                    Table(
                        table_data
                    )
                )

        return model