from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class ContentItem:
    order_index: int
    xml: str


@dataclass
class Run:
    """A single semantically-classified run of text within a paragraph or cell.

    Only content-owned emphasis (bold/italic/underline/hyperlink) is captured.
    Direct font/color/size formatting is intentionally NOT modeled here - the
    template style is the sole source of visual design during generation.
    """

    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    hyperlink_url: Optional[str] = None


@dataclass
class Paragraph(ContentItem):
    text: str
    style: Optional[str] = None
    runs: List[Run] = field(default_factory=list)
    keep_with_next: bool = False


@dataclass
class Heading(ContentItem):
    level: int
    text: str
    style: Optional[str] = None
    runs: List[Run] = field(default_factory=list)
    keep_with_next: bool = False


@dataclass
class Formula(ContentItem):
    """An equation/register-calculation line - content, never a heading.

    Detected by syntax (assignment operators, math functions) rather than by
    paragraph style/outline-level, so a source paragraph styled as a heading
    but containing a formula is reclassified here instead of becoming a
    numbered/TOC-eligible heading.
    """

    text: str
    style: Optional[str] = None
    runs: List[Run] = field(default_factory=list)
    keep_with_next: bool = False


@dataclass
class Cell:
    text: str
    runs: List[Run] = field(default_factory=list)
    row_span: int = 1
    col_span: int = 1
    is_header: bool = False
    raw_xml: str = ""


@dataclass
class Table(ContentItem):
    data: list
    cells: List[List[Cell]] = field(default_factory=list)


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