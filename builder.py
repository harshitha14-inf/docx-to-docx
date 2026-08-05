from copy import deepcopy
from zipfile import ZipFile

from lxml import etree

from models import FigureBlock


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


class Builder:

    def build(self, model, out_file):

        if not model.source_docx_path:
            raise ValueError(
                "DocumentModel.source_docx_path is required"
            )

        rebuilt_document = self._rebuild_document_xml(
            model
        )

        with ZipFile(
            model.source_docx_path,
            "r",
        ) as source_zip:

            with ZipFile(
                out_file,
                "w",
            ) as output_zip:

                for info in source_zip.infolist():

                    data = source_zip.read(
                        info.filename
                    )

                    if info.filename == "word/document.xml":
                        data = rebuilt_document

                    output_zip.writestr(
                        info,
                        data,
                    )

    def _rebuild_document_xml(
        self,
        model,
    ):

        with ZipFile(
            model.source_docx_path,
            "r",
        ) as source_zip:
            root = etree.fromstring(
                source_zip.read(
                    "word/document.xml"
                )
            )

        body = root.xpath(
            "./w:body",
            namespaces=NS,
        )[0]

        for child in list(body):
            body.remove(child)

        ordered_items = sorted(
            model.content,
            key=lambda item: item.order_index,
        )

        for item in ordered_items:

            for element in self._elements_for_item(
                item
            ):
                body.append(element)

        if model.body_sectpr_xml:
            body.append(
                etree.fromstring(
                    model.body_sectpr_xml.encode(
                        "utf-8"
                    )
                )
            )

        return etree.tostring(
            root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )

    def _elements_for_item(
        self,
        item,
    ):

        if isinstance(item, FigureBlock):

            elements = []

            if item.caption_position == "before":
                elements.extend(
                    self._elements_for_xml(
                        item.caption.xml
                    )
                )

            for image in item.images:
                elements.extend(
                    self._elements_for_xml(
                        image.xml
                    )
                )

            if item.caption_position == "after":
                elements.extend(
                    self._elements_for_xml(
                        item.caption.xml
                    )
                )

            return elements

        return self._elements_for_xml(
            item.xml
        )

    def _elements_for_xml(
        self,
        xml_text,
    ):

        element = etree.fromstring(
            xml_text.encode(
                "utf-8"
            )
        )

        return [
            deepcopy(element)
        ]

