from __future__ import annotations

from docxmdtpl.errors import DocumentStructureError
from docxmdtpl.errors import TemplateBindingError
from docxmdtpl.models import SectionNode
from docxmdtpl.rendering.template_profiles import PlaceholderProfile
from docxmdtpl.rendering.template_profiles import apply_paragraph_profile
from docxmdtpl.rendering.template_profiles import apply_run_profile


class HeadingRenderer:
    def render_heading(
        self,
        section: SectionNode,
        tpl: object,
        placeholder_profile: PlaceholderProfile | None = None,
    ) -> object:
        if section.level < 1 or section.level > 5:
            raise DocumentStructureError(
                "Heading fora do intervalo suportado pela V1.",
                level=section.level,
                title=section.title,
            )

        subdoc = self._create_subdoc(tpl)

        paragraph = self._add_heading_paragraph(
            subdoc=subdoc,
            level=section.level,
            placeholder_profile=placeholder_profile,
        )
        apply_paragraph_profile(
            paragraph,
            placeholder_profile,
            preserve_style=placeholder_profile is not None and placeholder_profile.paragraph.style_name is not None,
            preserve_indents=False,
        )

        run = paragraph.add_run(section.title)
        apply_run_profile(run, placeholder_profile)
        return subdoc

    def _create_subdoc(self, tpl: object) -> object:
        new_subdoc = getattr(tpl, "new_subdoc", None)
        if not callable(new_subdoc):
            raise TemplateBindingError("O template nao suporta criacao de subdocumentos.")

        return new_subdoc()

    @staticmethod
    def _heading_style_name(level: int) -> str:
        return f"Heading {level}"

    def _add_heading_paragraph(
        self,
        *,
        subdoc: object,
        level: int,
        placeholder_profile: PlaceholderProfile | None,
    ) -> object:
        preferred_style = None
        if placeholder_profile is not None:
            preferred_style = placeholder_profile.paragraph.style_name

        if preferred_style is not None:
            try:
                return subdoc.add_paragraph(style=preferred_style)
            except (KeyError, ValueError):
                pass

        try:
            return subdoc.add_paragraph(style=self._heading_style_name(level))
        except (KeyError, ValueError):
            return subdoc.add_paragraph()
