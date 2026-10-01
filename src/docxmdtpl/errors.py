from __future__ import annotations


class DocxMdTplError(Exception):
    """Erro base do dominio."""

    def __init__(self, message: str, **context: object) -> None:
        super().__init__(message)
        self.message = message
        self.context = context

    def __str__(self) -> str:
        if not self.context:
            return self.message
        rendered_context = ", ".join(
            f"{key}={value!r}" for key, value in sorted(self.context.items())
        )
        return f"{self.message} ({rendered_context})"


class MarkdownSourceError(DocxMdTplError):
    """Falha ao resolver a origem do Markdown."""


class FrontmatterError(DocxMdTplError):
    """Falha ao extrair ou validar frontmatter."""


class DocumentStructureError(DocxMdTplError):
    """Falha estrutural no conteudo Markdown."""


class UnsupportedMarkdownError(DocxMdTplError):
    """Falha por token ou bloco fora do escopo da V1."""


class ImageResolutionError(DocxMdTplError):
    """Falha ao localizar imagem referenciada no Markdown."""


class TemplateBindingError(DocxMdTplError):
    """Falha ao preparar ou renderizar o template .docx."""


class OutputPersistenceError(DocxMdTplError):
    """Falha ao persistir o arquivo final."""
