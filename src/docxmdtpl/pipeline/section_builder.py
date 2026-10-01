from __future__ import annotations

from docxmdtpl.errors import DocumentStructureError
from docxmdtpl.models import BlockLike, HeadingBlock, SectionNode


class SectionTreeBuilder:
    def build(self, blocks: list[BlockLike]) -> list[SectionNode]:
        sections: list[SectionNode] = []
        stack: list[SectionNode] = []

        for block in blocks:
            if isinstance(block, HeadingBlock):
                if block.level < 1 or block.level > 5:
                    raise DocumentStructureError(
                        "Heading fora do intervalo suportado pela V1.",
                        level=block.level,
                        title=block.title,
                    )

                section = SectionNode(level=block.level, title=block.title)

                while stack and stack[-1].level >= section.level:
                    stack.pop()

                if stack:
                    stack[-1].children.append(section)
                else:
                    sections.append(section)

                stack.append(section)
                continue

            if not stack:
                raise DocumentStructureError(
                    "Conteudo encontrado antes do primeiro heading.",
                    block_type=type(block).__name__,
                )

            stack[-1].blocks.append(block)

        return sections
