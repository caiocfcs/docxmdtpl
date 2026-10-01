from __future__ import annotations

from pathlib import Path

from docxmdtpl.errors import MarkdownSourceError
from docxmdtpl.models import ResolvedMarkdownSource


class SourceResolver:
    def resolve(self, markdown_content: str) -> ResolvedMarkdownSource:
        candidate_path = Path(markdown_content).expanduser()

        try:
            exists = candidate_path.exists()
        except OSError:
            return ResolvedMarkdownSource(
                original_input=markdown_content,
                markdown_text=markdown_content,
                source_path=None,
            )

        if exists:
            if not candidate_path.is_file():
                raise MarkdownSourceError(
                    "A origem do Markdown precisa apontar para um arquivo.",
                    markdown_content=markdown_content,
                )

            resolved_path = candidate_path.resolve()

            try:
                markdown_text = resolved_path.read_text(encoding="utf-8")
            except OSError as exc:
                raise MarkdownSourceError(
                    "Nao foi possivel ler o arquivo Markdown informado.",
                    markdown_content=markdown_content,
                    source_path=str(resolved_path),
                ) from exc

            return ResolvedMarkdownSource(
                original_input=markdown_content,
                markdown_text=markdown_text,
                source_path=str(resolved_path),
            )

        return ResolvedMarkdownSource(
            original_input=markdown_content,
            markdown_text=markdown_content,
            source_path=None,
        )
