from __future__ import annotations

from pathlib import Path
from urllib.parse import unquote

from docxmdtpl.errors import ImageResolutionError


class ImageResolver:
    def resolve(
        self,
        image_ref: str,
        path_imagens: str | None,
        source_dir: str | None,
    ) -> str:
        image_ref = image_ref.strip()
        if not image_ref:
            raise ImageResolutionError(
                "A referencia da imagem nao pode ser vazia.",
                image_ref=image_ref,
                path_imagens=path_imagens,
                source_dir=source_dir,
            )

        candidates: list[Path] = []
        for reference_variant in self._expand_reference_variants(image_ref):
            candidates.extend(self._build_candidates(reference_variant, path_imagens, source_dir))
        candidates = self._deduplicate_candidates(candidates)

        for candidate in candidates:
            if candidate.is_file():
                return str(candidate.resolve())

        raise ImageResolutionError(
            "Nao foi possivel localizar a imagem referenciada.",
            image_ref=image_ref,
            path_imagens=path_imagens,
            source_dir=source_dir,
        )

    @staticmethod
    def _expand_reference_variants(image_ref: str) -> list[str]:
        decoded_reference = unquote(image_ref)
        variants = [decoded_reference]

        if decoded_reference != image_ref:
            variants.append(image_ref)

        return variants

    def _build_candidates(
        self,
        image_ref: str,
        path_imagens: str | None,
        source_dir: str | None,
    ) -> list[Path]:
        reference_path = Path(image_ref).expanduser()
        candidates: list[Path] = []

        if reference_path.is_absolute():
            candidates.append(reference_path)
            return candidates

        if path_imagens:
            images_base = Path(path_imagens).expanduser()
            candidates.append(images_base / reference_path)

        if source_dir:
            source_base = Path(source_dir).expanduser()
            candidates.append(source_base / reference_path)

        candidates.append(reference_path)

        return self._deduplicate_candidates(candidates)

    @staticmethod
    def _deduplicate_candidates(candidates: list[Path]) -> list[Path]:
        seen: set[str] = set()
        unique_candidates: list[Path] = []

        for candidate in candidates:
            normalized = str(candidate)
            if normalized in seen:
                continue
            seen.add(normalized)
            unique_candidates.append(candidate)

        return unique_candidates
