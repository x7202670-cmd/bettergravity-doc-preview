from contextlib import redirect_stdout
import sys


def main():
    source = sys.stdin.buffer.read()
    if source.startswith(bytes.fromhex("d0cf11e0a1b11ae1")):
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
            preview = render_artifact(document.middle_json, "docx", assets=document.assets)
        except Exception as error:
            print(type(error).__name__, file=sys.stderr)
            return 2

    sys.stdout.buffer.write(preview.content)
    return 0


if __name__ == "__main__":
    sys.exit(main())
