"""Repair incomplete embedded CJK fonts in an in-memory preview copy."""

from collections import defaultdict
from functools import lru_cache
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
STYLES = "word/styles.xml"
RUN_FONTS = W + "rPr/" + W + "rFonts"
CJK_FONT_SLOTS = ("eastAsia", "ascii")
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


def font_slots(fonts):
    if fonts is None:
        return {}
    slots = {}
    for slot in CJK_FONT_SLOTS:
        if fonts.get(W + slot + "Theme"):
            # A theme value overrides a named font in the same slot.
            slots[slot] = None
        elif family := fonts.get(W + slot):
            slots[slot] = family
    return slots


def collect_font_usage(parts, embedded):
    styles = parts.get(STYLES)
    style_index = {}
    default_paragraph_style = None
    document_fonts = {}
    if styles is not None:
        document_fonts = font_slots(styles.find(W + "docDefaults/" + W + "rPrDefault/" + RUN_FONTS))
        for style in styles.findall(W + "style"):
            style_type = style.get(W + "type")
            style_id = style.get(W + "styleId")
            if style_id:
                style_index[style_type, style_id] = style
                if style_type == "paragraph" and style.get(W + "default") in ("1", "true", "on"):
                    default_paragraph_style = style_id

    @lru_cache(maxsize=256)
    def style_fonts(style_type, style_id):
        slots = {}
        visited = set()
        while style_id and style_id not in visited and len(slots) < len(CJK_FONT_SLOTS):
            style = style_index.get((style_type, style_id))
            if style is None:
                break
            visited.add(style_id)
            for slot, family in font_slots(style.find(RUN_FONTS)).items():
                slots.setdefault(slot, family)
            parent = style.find(W + "basedOn")
            style_id = parent.get(W + "val") if parent is not None else None
        return slots

    usage = defaultdict(set)
    for name, part in parts.items():
        if name in (STYLES, FONT_TABLE):
            continue
        for run in part.iter(W + "r"):
            required = cjk_characters("".join(t.text or "" for t in run.findall(W + "t")))
            if not required:
                continue
            paragraph = next(run.iterancestors(W + "p"), None)
            paragraph_style = paragraph.find(W + "pPr/" + W + "pStyle") if paragraph is not None else None
            paragraph_id = paragraph_style.get(W + "val") if paragraph_style is not None else None
            if ("paragraph", paragraph_id) not in style_index:
                paragraph_id = default_paragraph_style
            character_style = run.find(W + "rPr/" + W + "rStyle")
            character_id = character_style.get(W + "val") if character_style is not None else None
            slots = (document_fonts | style_fonts("paragraph", paragraph_id)
                     | style_fonts("character", character_id) | font_slots(run.find(RUN_FONTS)))
            family = slots.get("eastAsia", slots.get("ascii"))
            if family in embedded:
                usage[family].update(required)
    return usage


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
        for name in archive.namelist():
            if re.fullmatch(r"word/(?:document|header\d+|footer\d+|footnotes|endnotes|comments|styles)\.xml", name):
                parts[name] = read_xml(archive, name)
        usage = collect_font_usage(parts, embedded)

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
