from __future__ import annotations

import re


OBSIDIAN_IMAGE_PATTERN = re.compile(
    r"!\[\[(?P<reference>[^\]|]+?)(?:\|(?P<caption>[^\]]*?))?\]\]"
)


class ObsidianPreprocessor:
    def normalize(self, markdown_text: str) -> str:
        return OBSIDIAN_IMAGE_PATTERN.sub(self._replace_obsidian_image, markdown_text)

    @staticmethod
    def _replace_obsidian_image(match: re.Match[str]) -> str:
        reference = match.group("reference").strip()
        caption = match.group("caption")

        if not reference:
            return match.group(0)

        markdown_reference = ObsidianPreprocessor._to_markdown_destination(reference)

        if caption is None:
            return f"![]({markdown_reference})"

        cleaned_caption = caption.strip()
        if not cleaned_caption:
            return f"![]({markdown_reference})"

        escaped_caption = cleaned_caption.replace("[", r"\[").replace("]", r"\]")
        return f"![{escaped_caption}]({markdown_reference})"

    @staticmethod
    def _to_markdown_destination(reference: str) -> str:
        return f"<{reference}>"
