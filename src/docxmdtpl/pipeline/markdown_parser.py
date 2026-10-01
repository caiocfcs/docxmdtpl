from __future__ import annotations

from markdown_it import MarkdownIt
from markdown_it.token import Token


class MarkdownParser:
    def __init__(self) -> None:
        self._parser = MarkdownIt("commonmark", {"html": False}).enable("table")

    def parse(self, markdown_text: str) -> list[Token]:
        return self._parser.parse(markdown_text)
