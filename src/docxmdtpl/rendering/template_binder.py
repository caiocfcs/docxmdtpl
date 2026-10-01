from __future__ import annotations

from pathlib import Path

from docxtpl import DocxTemplate

from docxmdtpl.errors import OutputPersistenceError, TemplateBindingError
from docxmdtpl.models import DocumentModel, RenderedSection, SectionNode
from docxmdtpl.rendering.blocks import ContentRenderer
from docxmdtpl.rendering.headings import HeadingRenderer
from docxmdtpl.rendering.template_profiles import TemplateProfileExtractor
from docxmdtpl.rendering.template_profiles import TemplateRenderProfiles


class TemplateBinder:
    def __init__(
        self,
        heading_renderer: HeadingRenderer | None = None,
        content_renderer: ContentRenderer | None = None,
        profile_extractor: TemplateProfileExtractor | None = None,
    ) -> None:
        self._heading_renderer = heading_renderer or HeadingRenderer()
        self._content_renderer = content_renderer or ContentRenderer()
        self._profile_extractor = profile_extractor or TemplateProfileExtractor()

    def render(
        self,
        document: DocumentModel,
        template_path: str,
        output_path: str,
        path_imagens: str | None,
        source_dir: str | None,
    ) -> None:
        template = self._load_template(template_path)
        template_profiles = self._profile_extractor.extract(template)
        self._content_renderer.reset()
        context = {
            "metadata": document.metadata,
            "sections": [
                self._build_rendered_section(
                    section=section,
                    tpl=template,
                    path_imagens=path_imagens,
                    source_dir=source_dir,
                    template_profiles=template_profiles,
                )
                for section in document.sections
            ],
        }

        try:
            template.render(context)
        except Exception as exc:
            raise TemplateBindingError(
                "Falha ao renderizar o contexto no template .docx.",
                template_path=str(Path(template_path).expanduser()),
            ) from exc

        output = Path(output_path).expanduser()
        try:
            output.parent.mkdir(parents=True, exist_ok=True)
            template.save(str(output))
        except OSError as exc:
            raise OutputPersistenceError(
                "Falha ao salvar o arquivo .docx final.",
                output_path=str(output),
            ) from exc

    def _load_template(self, template_path: str) -> DocxTemplate:
        path = Path(template_path).expanduser()

        if path.suffix.lower() != ".docx":
            raise TemplateBindingError(
                "O template precisa ser um arquivo .docx.",
                template_path=str(path),
            )

        if not path.is_file():
            raise TemplateBindingError(
                "O template .docx informado nao foi encontrado.",
                template_path=str(path),
            )

        try:
            return DocxTemplate(str(path.resolve()))
        except Exception as exc:
            raise TemplateBindingError(
                "Nao foi possivel carregar o template .docx.",
                template_path=str(path),
            ) from exc

    def _build_rendered_section(
        self,
        *,
        section: SectionNode,
        tpl: DocxTemplate,
        path_imagens: str | None,
        source_dir: str | None,
        template_profiles: TemplateRenderProfiles,
    ) -> RenderedSection:
        return RenderedSection(
            level=section.level,
            title=section.title,
            heading_subdoc=self._heading_renderer.render_heading(
                section,
                tpl,
                placeholder_profile=template_profiles.resolve_heading(section.level),
            ),
            content_subdoc=self._content_renderer.render_content(
                section,
                tpl,
                path_imagens=path_imagens,
                source_dir=source_dir,
                placeholder_profile=template_profiles.resolve_content(section.level),
                figure_caption_template=template_profiles.resolve_figure_caption(section.level),
                quote_block_template=template_profiles.resolve_quote(section.level),
                code_block_template=template_profiles.resolve_code(section.level),
                markdown_table_template=template_profiles.resolve_table(section.level),
            ),
            children=[
                self._build_rendered_section(
                    section=child,
                    tpl=tpl,
                    path_imagens=path_imagens,
                    source_dir=source_dir,
                    template_profiles=template_profiles,
                )
                for child in section.children
            ],
        )
