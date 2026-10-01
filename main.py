from __future__ import annotations

import sys
from pathlib import Path


SRC_PATH = Path(__file__).resolve().parent / "src"
BASE_DIR = Path(__file__).resolve().parent

# Ajuste estes placeholders para os caminhos reais do seu ambiente.
TEMPLATE_PATH = BASE_DIR / "template.docx"
MARKDOWN_PATH = BASE_DIR / "conteudo.md"
OUTPUT_PATH = BASE_DIR / "saida" / "documento.docx"
IMAGES_PATH = BASE_DIR / "imagens"

if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from docxmdtpl import BuildTemplate, MarkdownToDocxTpl


def main() -> None:
    # Se nao existir template.docx, cria um template base com os placeholders esperados.
    if not TEMPLATE_PATH.is_file():
        BuildTemplate(TEMPLATE_PATH).build()

    if not MARKDOWN_PATH.is_file():
        raise FileNotFoundError(
            "Atualize MARKDOWN_PATH para um arquivo Markdown existente antes de executar o quickstart."
        )

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    MarkdownToDocxTpl(
        template_path=str(TEMPLATE_PATH),
        markdown_content=str(MARKDOWN_PATH),
        frontmatter=True,
        path_imagens=str(IMAGES_PATH),
    ).render(str(OUTPUT_PATH))

    print(f"DOCX gerado em: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
