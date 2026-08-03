from pathlib import Path
from zipfile import ZipFile
from shutil import copyfile

from docx import Document

from models import (
    DocumentModel,
    Paragraph,
    Table,
    Image
)


class Extractor:

    def __init__(self, docx_path):
        self.docx_path = docx_path

    def extract(self):

        document = Document(self.docx_path)

        model = DocumentModel()

        # ------------------------
        # Paragraphs
        # ------------------------

        for paragraph in document.paragraphs:

            text = paragraph.text.strip()

            if text:
                model.paragraphs.append(
                    Paragraph(text)
                )

        # ------------------------
        # Tables
        # ------------------------

        for table in document.tables:

            table_data = []

            for row in table.rows:

                row_data = []

                for cell in row.cells:
                    row_data.append(
                        cell.text.strip()
                    )

                table_data.append(
                    row_data
                )

            model.tables.append(
                Table(table_data)
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

        with ZipFile(self.docx_path) as z:

            media_files = [
                f for f in z.namelist()
                if f.startswith("word/media/")
            ]

            for media_file in media_files:

                image_name = Path(
                    media_file
                ).name

                output_file = (
                    media_dir /
                    image_name
                )

                with z.open(media_file) as src:
                    with open(
                        output_file,
                        "wb"
                    ) as dst:

                        dst.write(
                            src.read()
                        )

                model.images.append(
                    Image(
                        name=image_name,
                        path=str(output_file)
                    )
                )

        return model