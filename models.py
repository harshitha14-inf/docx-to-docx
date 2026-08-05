from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class ContentItem:
    order_index: int
    xml: str


@dataclass
class Paragraph(ContentItem):
    text: str
    style: Optional[str] = None


@dataclass
class Heading(ContentItem):
    level: int
    text: str
    style: Optional[str] = None


@dataclass
class Table(ContentItem):
    data: list


@dataclass
class ImageRef:
    name: str
    path: str
    relationship_id: str
    image_format: str
    document_position_index: int
    target: str


@dataclass
class Image(ContentItem):
    refs: List[ImageRef] = field(default_factory=list)


@dataclass
class Caption(ContentItem):
    text: str
    caption_type: str
    style: Optional[str] = None


@dataclass
class TextBox(ContentItem):
    texts: List[str] = field(default_factory=list)


@dataclass
class RawBlock(ContentItem):
    block_type: str


@dataclass
class FigureBlock(ContentItem):
    caption: Caption
    images: List[Image] = field(default_factory=list)
    caption_position: str = "before"


@dataclass
class Header:
    text: str
    style: Optional[str] = None


@dataclass
class Footer:
    text: str
    style: Optional[str] = None


@dataclass
class SectionMargins:
    left: int
    right: int
    top: int
    bottom: int
    header: int
    footer: int
    gutter: int


@dataclass
class SectionInfo:
    orientation: str
    page_width: int
    page_height: int
    start_type: str
    margins: SectionMargins


@dataclass
class DocumentModel:
    content: List[ContentItem] = field(default_factory=list)
    headers: List[Header] = field(default_factory=list)
    footers: List[Footer] = field(default_factory=list)
    sections: List[SectionInfo] = field(default_factory=list)

    source_docx_path: str = ""
    body_sectpr_xml: Optional[str] = None

    @property
    def paragraphs(self):
        return [
            item
            for item in self.content
            if isinstance(item, Paragraph)
        ]