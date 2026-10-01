from __future__ import annotations

from dataclasses import dataclass

from markdown_it.token import Token

from docxmdtpl.errors import UnsupportedMarkdownError
from docxmdtpl.models import (
    BlockLike,
    BoldSpan,
    BoldItalicSpan,
    CodeBlock,
    CodeSpan,
    HeadingBlock,
    HorizontalRuleBlock,
    ImageBlock,
    InlineSpan,
    ItalicSpan,
    LinkSpan,
    ListBlock,
    ListItem,
    ParagraphBlock,
    QuoteBlock,
    SpacerBlock,
    TableBlock,
    TableCell,
    TextSpan,
)


@dataclass(slots=True)
class _ParseResult:
    block: BlockLike
    next_index: int
    end_line: int | None = None


class BlockNormalizer:
    def normalize(self, tokens: list[Token]) -> list[BlockLike]:
        blocks, index = self._parse_block_sequence(tokens, 0)
        if index != len(tokens):
            raise UnsupportedMarkdownError(
                "Fluxo de blocos Markdown terminou em estado inesperado.",
                token_type=tokens[index].type,
                tag=tokens[index].tag,
            )
        return blocks

    def _parse_block_sequence(
        self,
        tokens: list[Token],
        start_index: int,
        *,
        stop_types: set[str] | None = None,
        allow_headings: bool = True,
    ) -> tuple[list[BlockLike], int]:
        blocks: list[BlockLike] = []
        index = start_index
        closing_types = stop_types or set()
        previous_end_line: int | None = None

        while index < len(tokens) and tokens[index].type not in closing_types:
            current_start_line = self._resolve_token_start_line(tokens[index])
            spacer_count = self._resolve_spacer_count(previous_end_line, current_start_line)
            if spacer_count > 0:
                blocks.append(SpacerBlock(line_breaks=spacer_count))

            result = self._parse_block(tokens, index, allow_headings=allow_headings)
            blocks.append(result.block)
            index = result.next_index
            previous_end_line = result.end_line

        return blocks, index

    def _parse_block(
        self,
        tokens: list[Token],
        start_index: int,
        *,
        allow_headings: bool,
    ) -> _ParseResult:
        token = tokens[start_index]

        if token.type == "heading_open":
            if not allow_headings:
                raise UnsupportedMarkdownError(
                    "Heading dentro de conteiner nao faz parte da V1.",
                    token_type=token.type,
                    tag=token.tag,
                )
            return self._parse_heading(tokens, start_index)

        if token.type == "paragraph_open":
            return self._parse_paragraph(tokens, start_index)

        if token.type in {"bullet_list_open", "ordered_list_open"}:
            return self._parse_list(tokens, start_index)

        if token.type == "blockquote_open":
            return self._parse_blockquote(tokens, start_index)

        if token.type == "table_open":
            return self._parse_table(tokens, start_index)

        if token.type in {"fence", "code_block"}:
            return self._parse_code_block(tokens, start_index)

        if token.type == "hr":
            return self._parse_horizontal_rule(tokens, start_index)

        raise UnsupportedMarkdownError(
            "Token Markdown fora da cobertura da V1.",
            token_type=token.type,
            tag=token.tag,
        )

    def _parse_heading(self, tokens: list[Token], start_index: int) -> _ParseResult:
        open_token = tokens[start_index]
        inline_token = self._require_token(tokens, start_index + 1, "inline")
        self._require_token(tokens, start_index + 2, "heading_close")

        return _ParseResult(
            block=HeadingBlock(level=int(open_token.tag.removeprefix("h")), title=inline_token.content),
            next_index=start_index + 3,
            end_line=self._resolve_token_end_line(open_token),
        )

    def _parse_paragraph(self, tokens: list[Token], start_index: int) -> _ParseResult:
        inline_token = self._require_token(tokens, start_index + 1, "inline")
        self._require_token(tokens, start_index + 2, "paragraph_close")

        children = inline_token.children or []
        image_block = self._extract_image_block(children)
        if image_block is not None:
            return _ParseResult(
                block=image_block,
                next_index=start_index + 3,
                end_line=self._resolve_token_end_line(tokens[start_index]),
            )

        return _ParseResult(
            block=ParagraphBlock(spans=self._parse_inline_spans(children)),
            next_index=start_index + 3,
            end_line=self._resolve_token_end_line(tokens[start_index]),
        )

    def _parse_list(self, tokens: list[Token], start_index: int) -> _ParseResult:
        open_token = tokens[start_index]
        closing_type = "bullet_list_close" if open_token.type == "bullet_list_open" else "ordered_list_close"
        index = start_index + 1
        items: list[ListItem] = []

        while index < len(tokens) and tokens[index].type != closing_type:
            self._require_token(tokens, index, "list_item_open")
            item, index = self._parse_list_item(tokens, index)
            items.append(item)

        self._require_token(tokens, index, closing_type)
        return _ParseResult(
            block=ListBlock(ordered=open_token.type == "ordered_list_open", items=items),
            next_index=index + 1,
            end_line=self._resolve_token_end_line(open_token),
        )

    def _parse_list_item(self, tokens: list[Token], start_index: int) -> tuple[ListItem, int]:
        blocks, index = self._parse_block_sequence(
            tokens,
            start_index + 1,
            stop_types={"list_item_close"},
            allow_headings=False,
        )
        spans: list[InlineSpan] = []
        children: list[BlockLike] = []

        for block in blocks:
            if isinstance(block, ParagraphBlock):
                if spans:
                    spans.append(TextSpan(text="\n"))
                spans.extend(block.spans)
                continue

            if isinstance(block, ImageBlock):
                raise UnsupportedMarkdownError(
                    "Listas da V1 nao suportam imagem como item isolado.",
                    block_type=type(block).__name__,
                )

            children.append(block)

        self._require_token(tokens, index, "list_item_close")
        return ListItem(spans=spans, children=children), index + 1

    def _parse_blockquote(self, tokens: list[Token], start_index: int) -> _ParseResult:
        blocks, index = self._parse_block_sequence(
            tokens,
            start_index + 1,
            stop_types={"blockquote_close"},
            allow_headings=False,
        )

        self._require_token(tokens, index, "blockquote_close")
        return _ParseResult(
            block=QuoteBlock(blocks=blocks),
            next_index=index + 1,
            end_line=self._resolve_token_end_line(tokens[start_index]),
        )

    def _parse_table(self, tokens: list[Token], start_index: int) -> _ParseResult:
        index = start_index + 1
        headers: list[TableCell] = []
        rows: list[list[TableCell]] = []

        self._require_token(tokens, index, "thead_open")
        index += 1

        while tokens[index].type != "thead_close":
            header_row, index = self._parse_table_row(tokens, index, cell_open_type="th_open", cell_close_type="th_close")
            headers.extend(header_row)

        self._require_token(tokens, index, "thead_close")
        index += 1

        self._require_token(tokens, index, "tbody_open")
        index += 1

        while tokens[index].type != "tbody_close":
            row, index = self._parse_table_row(tokens, index, cell_open_type="td_open", cell_close_type="td_close")
            rows.append(row)

        self._require_token(tokens, index, "tbody_close")
        index += 1
        self._require_token(tokens, index, "table_close")

        return _ParseResult(
            block=TableBlock(headers=headers, rows=rows),
            next_index=index + 1,
            end_line=self._resolve_token_end_line(tokens[start_index]),
        )

    def _parse_table_row(
        self,
        tokens: list[Token],
        start_index: int,
        *,
        cell_open_type: str,
        cell_close_type: str,
    ) -> tuple[list[TableCell], int]:
        self._require_token(tokens, start_index, "tr_open")
        index = start_index + 1
        cells: list[TableCell] = []

        while tokens[index].type != "tr_close":
            self._require_token(tokens, index, cell_open_type)
            inline_token = self._require_token(tokens, index + 1, "inline")
            self._require_token(tokens, index + 2, cell_close_type)
            cells.append(self._parse_table_cell(inline_token.children or []))
            index += 3

        self._require_token(tokens, index, "tr_close")
        return cells, index + 1

    def _parse_table_cell(self, children: list[Token]) -> TableCell:
        image_block = self._extract_image_block(children)
        if image_block is not None:
            return TableCell(blocks=[image_block])
        return TableCell(blocks=[ParagraphBlock(spans=self._parse_inline_spans(children))])

    def _parse_code_block(self, tokens: list[Token], start_index: int) -> _ParseResult:
        token = tokens[start_index]
        language = token.info.strip() or None
        return _ParseResult(
            block=CodeBlock(code=token.content, language=language),
            next_index=start_index + 1,
            end_line=self._resolve_token_end_line(token),
        )

    def _parse_horizontal_rule(self, tokens: list[Token], start_index: int) -> _ParseResult:
        self._require_token(tokens, start_index, "hr")
        return _ParseResult(
            block=HorizontalRuleBlock(),
            next_index=start_index + 1,
            end_line=self._resolve_token_end_line(tokens[start_index]),
        )

    def _extract_image_block(self, children: list[Token]) -> ImageBlock | None:
        meaningful_children = [
            child
            for child in children
            if child.type != "text" or child.content.strip()
        ]

        if len(meaningful_children) != 1 or meaningful_children[0].type != "image":
            return None

        image_token = meaningful_children[0]
        reference = (image_token.attrs or {}).get("src")
        caption = image_token.content or None

        if not reference:
            raise UnsupportedMarkdownError(
                "Imagem Markdown sem referencia valida.",
                token_type=image_token.type,
            )

        return ImageBlock(reference=reference, caption=caption)

    def _parse_inline_spans(self, children: list[Token]) -> list[InlineSpan]:
        spans, next_index = self._parse_inline_sequence(children, 0)
        if next_index != len(children):
            raise UnsupportedMarkdownError(
                "Sequencia inline Markdown inesperada.",
                token_type=children[next_index].type,
                markup=children[next_index].markup,
            )
        return spans

    def _parse_inline_sequence(
        self,
        children: list[Token],
        start_index: int,
        *,
        stop_at: str | None = None,
        strong_active: int = 0,
        emphasis_active: int = 0,
    ) -> tuple[list[InlineSpan], int]:
        spans: list[InlineSpan] = []
        index = start_index

        while index < len(children):
            child = children[index]

            if stop_at is not None and child.type == stop_at:
                return spans, index + 1

            if child.type == "text":
                if child.content:
                    spans.append(self._build_text_span(child.content, strong_active, emphasis_active))
                index += 1
                continue

            if child.type in {"softbreak", "hardbreak"}:
                spans.append(TextSpan(text="\n"))
                index += 1
                continue

            if child.type == "code_inline":
                spans.append(CodeSpan(text=child.content))
                index += 1
                continue

            if child.type == "strong_open":
                nested_spans, index = self._parse_inline_sequence(
                    children,
                    index + 1,
                    stop_at="strong_close",
                    strong_active=strong_active + 1,
                    emphasis_active=emphasis_active,
                )
                spans.extend(nested_spans)
                continue

            if child.type == "em_open":
                nested_spans, index = self._parse_inline_sequence(
                    children,
                    index + 1,
                    stop_at="em_close",
                    strong_active=strong_active,
                    emphasis_active=emphasis_active + 1,
                )
                spans.extend(nested_spans)
                continue

            if child.type == "link_open":
                nested_spans, index = self._parse_inline_sequence(
                    children,
                    index + 1,
                    stop_at="link_close",
                    strong_active=strong_active,
                    emphasis_active=emphasis_active,
                )
                link_target = self._extract_link_target(child)
                if link_target:
                    spans.append(
                        LinkSpan(
                            target=link_target,
                            spans=nested_spans or [TextSpan(text=link_target)],
                        )
                    )
                else:
                    spans.extend(nested_spans)
                continue

            if child.type == "image":
                raise UnsupportedMarkdownError(
                    "Imagens inline misturadas com texto nao fazem parte da V1.",
                    token_type=child.type,
                )

            if child.type.endswith("_close"):
                raise UnsupportedMarkdownError(
                    "Marcador inline Markdown inesperado.",
                    token_type=child.type,
                    markup=child.markup,
                )

            raise UnsupportedMarkdownError(
                "Inline Markdown fora da cobertura da V1.",
                token_type=child.type,
                markup=child.markup,
            )

        if stop_at is not None:
            raise UnsupportedMarkdownError(
                "Marcador inline Markdown sem fechamento.",
                expected_type=stop_at,
            )

        return spans, index

    @staticmethod
    def _build_text_span(text: str, strong_active: int, emphasis_active: int) -> InlineSpan:
        if strong_active and emphasis_active:
            return BoldItalicSpan(text=text)
        if strong_active:
            return BoldSpan(text=text)
        if emphasis_active:
            return ItalicSpan(text=text)
        return TextSpan(text=text)

    @staticmethod
    def _extract_link_target(token: Token) -> str | None:
        href = (token.attrs or {}).get("href")
        if href is None:
            return None
        normalized_href = href.strip()
        if normalized_href.startswith("#"):
            return None
        return normalized_href or None

    @staticmethod
    def _require_token(tokens: list[Token], index: int, expected_type: str) -> Token:
        if index >= len(tokens):
            raise UnsupportedMarkdownError(
                "Fluxo de tokens Markdown incompleto.",
                expected_type=expected_type,
            )

        token = tokens[index]
        if token.type != expected_type:
            raise UnsupportedMarkdownError(
                "Sequencia de tokens Markdown inesperada.",
                expected_type=expected_type,
                actual_type=token.type,
            )

        return token

    @staticmethod
    def _resolve_token_start_line(token: Token) -> int | None:
        if token.map is None:
            return None
        return token.map[0]

    @staticmethod
    def _resolve_token_end_line(token: Token) -> int | None:
        if token.map is None:
            return None
        return token.map[1]

    @staticmethod
    def _resolve_spacer_count(previous_end_line: int | None, current_start_line: int | None) -> int:
        if previous_end_line is None or current_start_line is None:
            return 0
        return max(0, current_start_line - previous_end_line - 1)
