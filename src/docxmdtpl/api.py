from __future__ import annotations

from docxmdtpl.models import DocumentModel
from docxmdtpl.pipeline.frontmatter import FrontmatterExtractor
from docxmdtpl.pipeline.markdown_parser import MarkdownParser
from docxmdtpl.pipeline.normalizer import BlockNormalizer
from docxmdtpl.pipeline.preprocessor import ObsidianPreprocessor
from docxmdtpl.pipeline.section_builder import SectionTreeBuilder
from docxmdtpl.pipeline.source_resolver import SourceResolver
from docxmdtpl.rendering.template_binder import TemplateBinder
from docxmdtpl.template_builder import BuildTemplate


class MarkdownToDocxTpl:
    def __init__(
        self,
        template_path: str,
        markdown_content: str,
        frontmatter: bool = True,
        path_imagens: str | None = None,
        source_dir: str | None = None,
    ) -> None:
        self.template_path = template_path
        self.markdown_content = markdown_content
        self.frontmatter = frontmatter
        self.path_imagens = path_imagens
        self._source_dir = source_dir

        self._source_resolver = SourceResolver()
        self._frontmatter_extractor = FrontmatterExtractor()
        self._preprocessor = ObsidianPreprocessor()
        self._markdown_parser = MarkdownParser()
        self._normalizer = BlockNormalizer()
        self._section_builder = SectionTreeBuilder()
        self._template_binder = TemplateBinder()

    def render(self, output_path: str) -> None:
        resolved_source = self._source_resolver.resolve(self.markdown_content)
        payload = self._frontmatter_extractor.extract(
            resolved_source.markdown_text,
            enabled=self.frontmatter,
        )
        normalized_markdown = self._preprocessor.normalize(payload.body)
        tokens = self._markdown_parser.parse(normalized_markdown)
        blocks = self._normalizer.normalize(tokens)
        sections = self._section_builder.build(blocks)

        document = DocumentModel(metadata=payload.metadata, sections=sections)
        effective_source_dir = self._source_dir or resolved_source.source_dir

        self._template_binder.render(
            document=document,
            template_path=self.template_path,
            output_path=output_path,
            path_imagens=self.path_imagens,
            source_dir=effective_source_dir,
        )
