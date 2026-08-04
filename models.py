from dataclasses import dataclass, field
from typing import List


@dataclass
class Paragraph:
    text: str


@dataclass
class Table:
    data: list


@dataclass
class Image:
    name: str
    path: str


@dataclass
class Caption:
    text: str
    caption_type: str
    # "image" or "table"


@dataclass
class Header:
    text: str


@dataclass
class Footer:
    text: str


@dataclass
class DocumentModel:

    # Ordered document content
    content: List = field(default_factory=list)

    # Headers & Footers
    headers: List[Header] = field(default_factory=list)
    footers: List[Footer] = field(default_factory=list)