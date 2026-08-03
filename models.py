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
class Header:
    name: str


@dataclass
class Footer:
    name: str


@dataclass
class DocumentModel:

    paragraphs: List[Paragraph] = field(default_factory=list)

    tables: List[Table] = field(default_factory=list)

    images: List[Image] = field(default_factory=list)

    headers: List[Header] = field(default_factory=list)

    footers: List[Footer] = field(default_factory=list)