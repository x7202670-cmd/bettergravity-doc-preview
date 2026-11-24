"""Repair incomplete embedded CJK fonts in an in-memory preview copy."""

from collections import defaultdict
from io import BytesIO
from pathlib import Path
import os
import posixpath
import re
import sys
from zipfile import ZipFile

from fontTools.ttLib import TTFont
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
FONT_TABLE = "word/fontTable.xml"
FONT_RELS = "word/_rels/fontTable.xml.rels"
FALLBACKS = {
    "仿宋_GB2312": ("FangSong", "仿宋"),
    "楷体_GB2312": ("KaiTi", "楷体"),
    "方正小标宋简体": ("STZhongsong", "华文中宋"),
}


def read_xml(archive, name):
    if archive.getinfo(name).file_size > 32 * 1024 * 1024:
        raise ValueError("DOCX XML part too large")
    return etree.fromstring(archive.read(name), etree.XMLParser(resolve_entities=False, no_network=True))


def cjk_characters(text):
    return {ord(ch) for ch in text if "\u3400" <= ch <= "\u9fff" or "\U00020000" <= ch <= "\U000323af"}


def installed_fonts():
    if sys.platform != "win32":
        return {}
    import winreg

    fonts = {}
    key_name = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive, key_name) as key:
                for index in range(winreg.QueryInfoKey(key)[1]):
                    name, filename, _ = winreg.EnumValue(key, index)
                    if not isinstance(filename, str):
                        continue
                    path = Path(os.path.expandvars(filename))
                    if not path.is_absolute():
                        directory = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
                        if hive == winreg.HKEY_CURRENT_USER:
                            directory = Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Windows" / "Fonts"
                        path = directory / path
                    family = re.sub(r"\s*\([^)]*\)$", "", name)
                    for alias in family.split(" & "):
                        fonts[alias.casefold()] = path
        except OSError:
            continue
    return fonts


def font_coverage(source):
    with TTFont(source, fontNumber=0, lazy=True) as font:
        return set(font.getBestCmap() or {})


def embedded_coverage(archive, ref, relationships):
    target = relationships.get(ref.get(R + "id"))
    key = ref.get(W + "fontKey")
    if not target or not key:
        return None
    name = posixpath.normpath(posixpath.join("word", target))
    if not name.startswith("word/") or archive.getinfo(name).file_size > 32 * 1024 * 1024:
        return None
    data = bytearray(archive.read(name))
    key_bytes = bytes.fromhex(key.strip("{}").replace("-", ""))[::-1]
    if len(key_bytes) != 16 or len(data) < 32:
        return None
    for index in range(32):
        data[index] ^= key_bytes[index % 16]
    return font_coverage(BytesIO(data))


def normalize(source):
    with ZipFile(BytesIO(source)) as archive:
        # Avoid expanding unusually large packages just to generate an optional preview.
        if sum(info.file_size for info in archive.infolist()) > 512 * 1024 * 1024:
            return source
        if FONT_TABLE not in archive.namelist() or FONT_RELS not in archive.namelist():
            return source
        table = read_xml(archive, FONT_TABLE)
        embedded = {
            font.get(W + "name"): font for font in table.findall(W + "font")
            if any(child.tag.startswith(W + "embed") for child in font)
        }
        if not embedded:
            return source

        parts = {FONT_TABLE: table}
        usage = defaultdict(set)
        for name in archive.namelist():
            if not re.fullmatch(r"word/(?:document|header\d+|footer\d+|footnotes|endnotes|comments|styles)\.xml", name):
                continue
            root = parts[name] = read_xml(archive, name)
            for run in root.iter(W + "r"):
                fonts = run.find(W + "rPr/" + W + "rFonts")
                if fonts is None:
                    continue
                family = fonts.get(W + "eastAsia") or fonts.get(W + "ascii")
                if family in embedded:
                    usage[family].update(cjk_characters("".join(t.text or "" for t in run.iter(W + "t"))))

        relationships = {rel.get("Id"): rel.get("Target") for rel in read_xml(archive, FONT_RELS)
                         if rel.get("TargetMode") != "External"}
        local_fonts = installed_fonts()
        replacements = {}
        for family, required in usage.items():
            if not required:
                continue
            try:
                coverages = [embedded_coverage(archive, ref, relationships) for ref in embedded[family]
                             if ref.tag.startswith(W + "embed")]
            except Exception:
                continue
            if not coverages or not any(coverage is not None and required - coverage for coverage in coverages):
                continue
            for candidate in (family, *FALLBACKS.get(family, ())):
                path = local_fonts.get(candidate.casefold())
                try:
                    if path is not None and required <= font_coverage(path):
                        replacements[family] = candidate
                        break
                except Exception:
                    continue

        if not replacements:
            return source

        for root in parts.values():
            for fonts in root.iter(W + "rFonts"):
                for attribute in ("ascii", "eastAsia", "hAnsi", "cs"):
                    key = W + attribute
                    if fonts.get(key) in replacements:
                        fonts.set(key, replacements[fonts.get(key)])

        # The repaired preview uses complete installed fonts instead of incomplete subsets.
        for font in list(table.findall(W + "font")):
            family = font.get(W + "name")
            if family in replacements and replacements[family] != family:
                table.remove(font)
            elif family in replacements.values():
                for child in list(font):
                    if child.tag.startswith(W + "embed"):
                        font.remove(child)

        output = BytesIO()
        with ZipFile(output, "w") as repaired:
            for info in archive.infolist():
                data = archive.read(info)
                if info.filename in parts:
                    data = etree.tostring(parts[info.filename], encoding="UTF-8", xml_declaration=True)
                repaired.writestr(info, data)
        return output.getvalue()


if __name__ == "__main__":
    sys.stdout.buffer.write(normalize(sys.stdin.buffer.read()))
