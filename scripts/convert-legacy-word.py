from contextlib import redirect_stdout
from io import BytesIO
import sys


def render_slide_pages(document, render_artifact):
    from pypdfium2 import PdfDocument

    with PdfDocument.new() as result:
        for slide in document.middle_json.pages:
            content = document.middle_json.model_copy(update={"pages": [slide]})
            preview = render_artifact(content, "pdf", assets=document.assets)
            with PdfDocument(preview.content) as pages:
                result.import_pages(pages)
        output = BytesIO()
        result.save(output)
        return output.getvalue()


def main():
    source = sys.stdin.buffer.read()
    output_format = "docx"
    if len(sys.argv) > 1 and sys.argv[1] in ("ppt", "pptx"):
        suffix = sys.argv[1]
        output_format = "pdf"
    elif source.startswith(bytes.fromhex("d0cf11e0a1b11ae1")):
        suffix = "doc"
    elif source.startswith(b"{\\rtf"):
        suffix = "rtf"
    elif source.startswith(b"PK\x03\x04"):
        suffix = "docx"
    else:
        return 2

    with redirect_stdout(sys.stderr):
        from docvortex import parse, render_artifact

        try:
            document = parse(source, file_suffix=suffix)
            if output_format == "pdf":
                preview = render_slide_pages(document, render_artifact)
            else:
                preview = render_artifact(document.middle_json, "docx", assets=document.assets).content
        except Exception as error:
            print(type(error).__name__, file=sys.stderr)
            return 2

    sys.stdout.buffer.write(preview)
    return 0


if __name__ == "__main__":
    sys.exit(main())
