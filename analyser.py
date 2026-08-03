from pathlib import Path
from zipfile import ZipFile, BadZipFile
from lxml import etree
import json


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}


class DocxAnalyzer:

    def __init__(self, docx_path):
        self.docx_path = Path(docx_path)

    def _read_xml(self, zip_file, xml_path):
        try:
            xml_content = zip_file.read(xml_path)
            return etree.fromstring(xml_content)
        except Exception:
            return None

    def analyze(self):

        report = {
            "file_name": self.docx_path.name,
            "file_size_kb": round(
                self.docx_path.stat().st_size / 1024,
                2
            ),

            "paragraphs_total": 0,
            "paragraphs_non_empty": 0,

            "tables": 0,

            "images": 0,
            "image_extensions": [],

            "drawings_total": 0,
            "image_drawings": 0,
            "non_image_drawings": 0,

            "inline_images": 0,
            "floating_images": 0,

            "textboxes": 0,

            "headers": 0,
            "footers": 0,

            "charts": 0,
            "ole_objects": 0,
            "vml_shapes": 0,

            "section_breaks": 0,

            "media_files": []
        }

        if not self.docx_path.exists():
            raise FileNotFoundError(
                f"File not found: {self.docx_path}"
            )

        with ZipFile(self.docx_path, "r") as zip_file:

            files = zip_file.namelist()

            # --------------------------------------------------
            # MEDIA FILES
            # --------------------------------------------------

            media = [
                f
                for f in files
                if f.startswith("word/media/")
            ]

            report["media_files"] = media
            report["images"] = len(media)

            extensions = set()

            for media_file in media:

                ext = Path(media_file).suffix.lower()

                if ext:
                    extensions.add(ext)

            report["image_extensions"] = sorted(
                list(extensions)
            )

            # --------------------------------------------------
            # HEADERS
            # --------------------------------------------------

            report["headers"] = len(
                [
                    f
                    for f in files
                    if f.startswith("word/header")
                    and f.endswith(".xml")
                ]
            )

            # --------------------------------------------------
            # FOOTERS
            # --------------------------------------------------

            report["footers"] = len(
                [
                    f
                    for f in files
                    if f.startswith("word/footer")
                    and f.endswith(".xml")
                ]
            )

            # --------------------------------------------------
            # MAIN DOCUMENT XML
            # --------------------------------------------------

            root = self._read_xml(
                zip_file,
                "word/document.xml"
            )

            if root is None:
                return report

            # --------------------------------------------------
            # PARAGRAPHS
            # --------------------------------------------------

            paragraphs = root.xpath(
                ".//w:p",
                namespaces=NS
            )

            report["paragraphs_total"] = len(
                paragraphs
            )

            non_empty = 0

            for paragraph in paragraphs:

                text = "".join(
                    paragraph.xpath(
                        ".//w:t/text()",
                        namespaces=NS
                    )
                ).strip()

                if text:
                    non_empty += 1

            report["paragraphs_non_empty"] = non_empty

            # --------------------------------------------------
            # TABLES
            # --------------------------------------------------

            report["tables"] = len(
                root.xpath(
                    ".//w:tbl",
                    namespaces=NS
                )
            )

            # --------------------------------------------------
            # DRAWINGS
            # --------------------------------------------------

            drawings = root.xpath(
                ".//w:drawing",
                namespaces=NS
            )

            report["drawings_total"] = len(
                drawings
            )

            image_drawings = 0
            non_image_drawings = 0

            inline_images = 0
            floating_images = 0

            for drawing in drawings:

                has_image = bool(
                    drawing.xpath(
                        ".//*[local-name()='blip']"
                    )
                )

                if has_image:

                    image_drawings += 1

                    if drawing.xpath(
                        ".//*[local-name()='inline']"
                    ):
                        inline_images += 1

                    if drawing.xpath(
                        ".//*[local-name()='anchor']"
                    ):
                        floating_images += 1

                else:
                    non_image_drawings += 1

            report["image_drawings"] = image_drawings
            report["non_image_drawings"] = non_image_drawings

            report["inline_images"] = inline_images
            report["floating_images"] = floating_images

            # --------------------------------------------------
            # TEXTBOXES
            # --------------------------------------------------

            report["textboxes"] = len(
                root.xpath(
                    ".//*[local-name()='txbxContent']"
                )
            )

            # --------------------------------------------------
            # OLE OBJECTS
            # --------------------------------------------------

            report["ole_objects"] = len(
                root.xpath(
                    ".//*[local-name()='OLEObject']"
                )
            )

            # --------------------------------------------------
            # VML SHAPES
            # --------------------------------------------------

            report["vml_shapes"] = len(
                root.xpath(
                    ".//*[local-name()='shape']"
                )
            )

            # --------------------------------------------------
            # CHARTS
            # --------------------------------------------------

            report["charts"] = len(
                root.xpath(
                    ".//*[local-name()='chart']"
                )
            )

            # --------------------------------------------------
            # SECTION BREAKS
            # --------------------------------------------------

            report["section_breaks"] = len(
                root.xpath(
                    ".//w:sectPr",
                    namespaces=NS
                )
            )

        return report


if __name__ == "__main__":

    import argparse

    try:

        parser = argparse.ArgumentParser()

        parser.add_argument(
            "docx",
            help="Path to DOCX file"
        )

        args = parser.parse_args()

        analyzer = DocxAnalyzer(
            args.docx
        )

        report = analyzer.analyze()

        print(
            json.dumps(
                report,
                indent=4
            )
        )

    except BadZipFile:
        print(
            "ERROR: Invalid DOCX file."
        )

    except FileNotFoundError as e:
        print(
            f"ERROR: {e}"
        )

    except Exception as e:
        print(
            f"ERROR: {repr(e)}"
        )