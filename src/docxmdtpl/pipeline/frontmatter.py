from __future__ import annotations

import frontmatter

from docxmdtpl.errors import FrontmatterError
from docxmdtpl.models import MarkdownPayload


class FrontmatterExtractor:
    def extract(self, markdown_text: str, enabled: bool) -> MarkdownPayload:
        if not enabled:
            return MarkdownPayload(metadata={}, body=markdown_text)

        try:
            post = frontmatter.loads(markdown_text)
        except Exception as exc:
            raise FrontmatterError(
                "Frontmatter YAML invalido.",
                detail=str(exc),
            ) from exc

        return MarkdownPayload(metadata=dict(post.metadata), body=post.content)
