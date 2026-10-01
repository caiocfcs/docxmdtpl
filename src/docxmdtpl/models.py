from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import TypeAlias


@dataclass(slots=True)
class ResolvedMarkdownSource:
    original_input: str
    markdown_text: str
    source_path: str | None

    @property
    def source_dir(self) -> str | None:
        if self.source_path is None:
            return None
        return str(Path(self.source_path).parent)


@dataclass(slots=True)
class MarkdownPayload:
    metadata: dict[str, object]
    body: str


@dataclass(slots=True)
class TextSpan:
    text: str


@dataclass(slots=True)
class BoldSpan:
    text: str


@dataclass(slots=True)
class ItalicSpan:
    text: str


@dataclass(slots=True)
class BoldItalicSpan:
    text: str


@dataclass(slots=True)
class CodeSpan:
    text: str


@dataclass(slots=True)
class LinkSpan:
    target: str
    spans: list["InlineSpan"]


InlineSpan: TypeAlias = TextSpan | BoldSpan | ItalicSpan | BoldItalicSpan | CodeSpan | LinkSpan


@dataclass(slots=True)
class HeadingBlock:
    level: int
    title: str


@dataclass(slots=True)
class ParagraphBlock:
    spans: list[InlineSpan]


@dataclass(slots=True)
class ListItem:
    spans: list[InlineSpan]
    children: list["BlockLike"] = field(default_factory=list)


@dataclass(slots=True)
class ListBlock:
    ordered: bool
    items: list[ListItem]


@dataclass(slots=True)
class TableCell:
    blocks: list["BlockLike"]


@dataclass(slots=True)
class TableBlock:
    headers: list[TableCell]
    rows: list[list[TableCell]]


@dataclass(slots=True, init=False)
class QuoteBlock:
    blocks: list["BlockLike"]

    def __init__(
        self,
        blocks: list["BlockLike"] | None = None,
        *,
        paragraphs: list[list[InlineSpan]] | None = None,
    ) -> None:
        if blocks is not None and paragraphs is not None:
            raise ValueError("QuoteBlock aceita `blocks` ou `paragraphs`, nao ambos.")

        if blocks is None:
            blocks = [ParagraphBlock(spans=spans) for spans in (paragraphs or [])]

        self.blocks = blocks

    @property
    def paragraphs(self) -> list[list[InlineSpan]]:
        return [
            block.spans
            for block in self.blocks
            if isinstance(block, ParagraphBlock)
        ]


@dataclass(slots=True)
class CodeBlock:
    code: str
    language: str | None = None


@dataclass(slots=True)
class ImageBlock:
    reference: str
    caption: str | None = None


@dataclass(slots=True)
class HorizontalRuleBlock:
    pass


@dataclass(slots=True)
class SpacerBlock:
    line_breaks: int = 1


BlockLike: TypeAlias = (
    HeadingBlock
    | ParagraphBlock
    | ListBlock
    | TableBlock
    | QuoteBlock
    | CodeBlock
    | ImageBlock
    | HorizontalRuleBlock
    | SpacerBlock
)


@dataclass(slots=True)
class SectionNode:
    level: int
    title: str
    blocks: list[BlockLike] = field(default_factory=list)
    children: list["SectionNode"] = field(default_factory=list)


@dataclass(slots=True)
class DocumentModel:
    metadata: dict[str, object]
    sections: list[SectionNode]


@dataclass(slots=True)
class RenderedSection:
    level: int
    title: str
    heading_subdoc: object
    content_subdoc: object
    children: list["RenderedSection"] = field(default_factory=list)
