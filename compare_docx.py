"""
compare_docx.py

A simple point-and-click window for comparing two Word (.docx) files.
Pick the old file, pick the new file, choose where to save the report,
and click Compare. No command line needed. Each file can be a Word
(.docx) file or a PDF, so Word-vs-Word, PDF-vs-PDF and Word-vs-PDF all work.
Old Word 97-2003 (.doc) files are converted to .docx first, using
Microsoft Word (on Windows, with pywin32) or LibreOffice. HTML pages
(.html/.htm, and Word's single-file .mht pages) can be compared too.

Text is compared word-by-word, and images are shown inline and matched
by their actual content. The window is built with tkinter, which comes
built into Python. Extra packages: python-docx (always), pillow
(recommended, for comparing figures) and pymupdf (only for PDF files).

Usage:
    python compare_docx.py
"""

import base64
import codecs
import difflib
import email
import email.policy
import hashlib
import html
import io
import os
import pathlib
import re
import shutil
import string
import subprocess
import tempfile
import threading
import webbrowser
import zipfile
from collections import Counter
from html.parser import HTMLParser
from urllib.parse import unquote, urljoin
from urllib.request import url2pathname

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from docx import Document
from docx.oxml.ns import qn

try:  # Optional: lets pictures be compared by what they look like.
    from PIL import Image, ImageChops
except ImportError:
    Image = None

try:  # Optional: needed only for comparing PDF files.
    import pymupdf
except ImportError:
    try:
        import fitz as pymupdf  # older PyMuPDF versions
    except ImportError:
        pymupdf = None


# ===========================================================================
# Core comparison logic
# ===========================================================================

W_P = qn("w:p")
W_R = qn("w:r")
W_T = qn("w:t")
W_TAB = qn("w:tab")
W_BREAKS = (qn("w:br"), qn("w:cr"))
# Word stores some content (e.g. text boxes) twice: a modern version and a
# legacy "fallback" copy for old versions of Word. We skip the fallback copy
# so that content isn't reported twice.
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"

# Two paragraphs are treated as "the same paragraph, edited" when at least
# this share of their words match. Below it they count as removed + added.
SIMILARITY_THRESHOLD = 0.5
# A requirement whose number changed is still recognised as the same
# requirement when at least this share of its content matches.
RENUMBER_THRESHOLD = 0.6
# Paragraphs starting with this text mark the beginning of a requirement.
DEFAULT_REQUIREMENT_PREFIX = "Requirement Number"
# Largest changed block (old paragraphs x new paragraphs) that gets the
# careful similarity alignment; bigger blocks use a faster simple pairing.
MAX_ALIGN_CELLS = 10000


def _walk(element, stay_in_paragraph=False):
    """
    Yield every element below `element`, in document order, skipping legacy
    fallback copies. With stay_in_paragraph=True, paragraphs nested inside
    (e.g. the contents of a text box) are skipped too, because they are
    reported as paragraphs of their own.
    """
    for child in element:
        if not isinstance(child.tag, str):
            continue  # XML comments / processing instructions
        if child.tag == MC_FALLBACK:
            continue
        if stay_in_paragraph and child.tag == W_P:
            continue
        yield child
        yield from _walk(child, stay_in_paragraph)


def image_fingerprint(data):
    """
    Return a fingerprint that is equal for two pictures that look the same.

    With Pillow installed, the picture is decoded and its pixels are
    fingerprinted, so a figure that Word merely re-saved (different
    compression, metadata or file type, identical pixels) still counts as
    unchanged. Without Pillow, or for formats it can't read (EMF/WMF), the
    raw file bytes are used, which treats any re-saved file as different.
    """
    if Image is not None:
        try:
            with Image.open(io.BytesIO(data)) as im:
                im = im.convert("RGBA")
                pixels = im.tobytes()
                return hashlib.md5(repr(im.size).encode() + pixels).hexdigest()
        except Exception:
            pass
    return hashlib.md5(data).hexdigest()


def heading_styles(doc):
    """
    Return the set of paragraph style IDs that are titles/headings in this
    document: Word's built-in "Title", "Subtitle" and "Heading 1-9" styles,
    any style given an outline level, and any style based on one of those.
    (Word keeps the built-in style *names* in English in every language
    version, so this also works for non-English documents.)
    """
    info = {}
    for style in doc.styles.element.iter(qn("w:style")):
        style_id = style.get(qn("w:styleId"))
        name_el = style.find(qn("w:name"))
        based_el = style.find(qn("w:basedOn"))
        name = (name_el.get(qn("w:val")) or "").lower() if name_el is not None else ""
        based_on = based_el.get(qn("w:val")) if based_el is not None else None
        own = (name in ("title", "subtitle") or name.startswith("heading ")
               or _has_outline_level(style.find(qn("w:pPr"))))
        info[style_id] = (own, based_on)

    result = set()
    for style_id in info:
        current, hops = style_id, 0
        while current in info and hops < 20:
            own, based_on = info[current]
            if own:
                result.add(style_id)
                break
            current, hops = based_on, hops + 1
    return result


def _has_outline_level(ppr):
    """True if these paragraph properties set an outline level (0-8 = heading levels)."""
    if ppr is None:
        return False
    lvl = ppr.find(qn("w:outlineLvl"))
    if lvl is None:
        return False
    try:
        return int(lvl.get(qn("w:val"))) < 9
    except (TypeError, ValueError):
        return False


def is_heading(p_elm, heading_style_ids):
    """Is this paragraph a title/heading (by its style or its outline level)?"""
    ppr = p_elm.find(qn("w:pPr"))
    if ppr is None:
        return False
    if _has_outline_level(ppr):
        return True
    style = ppr.find(qn("w:pStyle"))
    return style is not None and style.get(qn("w:val")) in heading_style_ids


def extract_items(path):
    """
    Walk through a .docx file's raw XML in document order and return a list
    of paragraph items, one per non-empty paragraph:

        {"text": "...", "heading": True/False,
         "images": [{"filename": ..., "partname": ..., "bytes": ..., "hash": ...}, ...]}

    "heading" is True for titles (paragraphs using a Word heading/title style).

    Scanning the raw XML (rather than only python-docx's top-level paragraph
    list) means content is found no matter how it's nested — inside tables,
    nested tables, text boxes, or content controls (structured document tags).

    Two kinds of image references are recognized:
      - Modern images:  <a:blip r:embed="..."/>
      - Older / pasted-in images that still use the legacy VML format:
        <v:imagedata r:id="..."/>

    Tabs and line breaks inside a paragraph are kept as whitespace so the
    words on either side of them stay separate.

    Paragraphs with no text and no images are skipped.
    """
    doc = Document(path)
    body = doc.element.body
    items = []
    fingerprints = {}
    heading_ids = heading_styles(doc)

    for p_elm in _walk(body):
        if p_elm.tag != W_P:
            continue

        text_parts = []
        images = []
        for el in _walk(p_elm, stay_in_paragraph=True):
            tag = el.tag
            if tag == W_T:
                text_parts.append(el.text or "")
                continue
            if tag == W_TAB:
                # <w:tab> also appears in tab-stop definitions; only the
                # ones directly inside a run are real tab characters.
                parent = el.getparent()
                if parent is not None and parent.tag == W_R:
                    text_parts.append("\t")
                continue
            if tag in W_BREAKS:
                text_parts.append("\n")
                continue

            local_tag = tag.split("}")[-1]
            if local_tag == "blip":
                rid = el.get(qn("r:embed"))
            elif local_tag == "imagedata":
                rid = el.get(qn("r:id"))
            else:
                continue
            if not rid:
                continue
            try:
                part = doc.part.related_parts[rid]
            except KeyError:
                continue
            data = part.blob
            partname = str(part.partname)
            if partname not in fingerprints:  # a reused picture is decoded once
                fingerprints[partname] = image_fingerprint(data)
            images.append({
                "filename": partname.split("/")[-1],
                "partname": partname,
                "bytes": data,
                "hash": fingerprints[partname],
            })

        text = "".join(text_parts).strip()
        if not text and not images:
            continue  # skip fully empty paragraphs

        items.append({"text": text, "images": images,
                      "heading": bool(text) and is_heading(p_elm, heading_ids)})

    return items


def list_media_files(path):
    """Return the set of image files physically stored inside the .docx (word/media/)."""
    with zipfile.ZipFile(path) as z:
        return {n for n in z.namelist() if n.startswith("word/media/") and not n.endswith("/")}


# ===========================================================================
# PDF files
# ===========================================================================
#
# A PDF doesn't store paragraphs, headings or "this is a header" -- only
# positioned lines of text and pictures. The functions below rebuild
# paragraphs from the line layout, recognise titles by their larger font,
# and drop page headers/footers (text repeated at the top/bottom of pages,
# and page numbers) so they don't show up as differences.

# Top/bottom share of each page that is searched for headers and footers.
PDF_MARGIN_SHARE = 0.09
# A paragraph whose font is at least this many points larger than the
# body text is treated as a title.
PDF_HEADING_EXTRA_PT = 0.9

# A list/heading number or bullet ("1.", "3.2", "a)", "iv.", "•", ...).
_MARKER = (r"(?:[•●▪■◦○·◆►✓\-–—*\uf000-\uf0ff]|\(?\d+(?:\.\d+)*[.)]?"
           r"|\(?[A-Za-z][.)]|\(?[ivxlcIVXLC]+[.)])")
_MARKER_ONLY = re.compile(_MARKER)
# A PDF line that begins with such a marker starts a new paragraph.
_LIST_START = re.compile(r"^\s*" + _MARKER + r"(?:\s|$)")
_PAGE_NUMBER = re.compile(r"(?:-\s*)?(?:page|p\.|sayfa|seite)?\s*#\s*(?:(?:of|/|von|de|-)\s*#)?(?:\s*-)?")
_DOT_LEADER = re.compile(r"\s*(?:\.\s*){4,}|\s*(?:…\s*){2,}")
_TERMINAL = (".", ":", ";", "!", "?")


def is_pdf(path):
    """True if the file is a PDF (checked by its content, not just its name)."""
    with open(path, "rb") as f:
        return f.read(5) == b"%PDF-"


def _pdf_image(block, page_number, index):
    """Turn a PDF image block into an image dict like the ones from Word files."""
    data, ext = block["image"], (block.get("ext") or "png").lower()
    mask = block.get("mask")
    if Image is not None and (mask or ext not in ("png", "jpeg", "jpg", "gif", "bmp")):
        # Re-apply the transparency mask (stored separately in PDFs), and
        # convert formats browsers can't show (JPEG 2000, JBIG2, ...) to PNG.
        try:
            with Image.open(io.BytesIO(data)) as im:
                im.load()
                if mask:
                    with Image.open(io.BytesIO(mask)) as m:
                        alpha = m.convert("L")
                        if alpha.size != im.size:
                            alpha = alpha.resize(im.size)
                    im = im.convert("RGB")
                    im.putalpha(alpha)
                out = io.BytesIO()
                im.save(out, "PNG")
                data, ext = out.getvalue(), "png"
        except Exception:
            pass
    name = f"page{page_number + 1}-image{index}.{ext}"
    return {"filename": name, "partname": name, "bytes": data,
            "hash": image_fingerprint(data)}


def extract_items_pdf(path, split_before=""):
    """
    Read a PDF and return paragraph items in the same shape as extract_items():
        {"text": ..., "heading": bool, "images": [...]}

    Pictures become items of their own (text ""), in reading order.
    `split_before` (e.g. "Requirement Number") always starts a new paragraph,
    so a requirement is never glued onto the end of the paragraph above it.
    """
    if pymupdf is None:
        raise RuntimeError(
            "Comparing PDF files needs the PyMuPDF package.\n"
            "Install it with:  pip install pymupdf"
        )
    name = os.path.basename(path)
    doc = pymupdf.open(path)
    try:
        if doc.needs_pass:
            raise ValueError(f"{name} is password-protected. Remove the password and try again.")

        flags = (pymupdf.TEXTFLAGS_DICT | pymupdf.TEXT_PRESERVE_IMAGES) & ~pymupdf.TEXT_PRESERVE_LIGATURES
        pages = []
        for page in doc:
            height = page.rect.height
            elements = []
            for b in page.get_text("dict", sort=True, flags=flags)["blocks"]:
                if b["type"] == 1:
                    x0, y0, x1, y1 = b["bbox"]
                    if x1 - x0 >= 3 and y1 - y0 >= 3 and b.get("image"):
                        elements.append({"type": "img", "bbox": b["bbox"], "block": b,
                                         "key": "i:" + hashlib.md5(b["image"]).hexdigest()})
                    continue
                for ln in b["lines"]:
                    all_spans = ln["spans"]
                    spans = [s for s in all_spans if s["text"].strip()]
                    if not spans:
                        continue
                    text = "".join(s["text"] for s in all_spans).strip()
                    size = max(spans, key=lambda s: len(s["text"].strip()))["size"]
                    element = {
                        "type": "line", "text": text, "bbox": ln["bbox"], "size": size,
                        "block": b["number"], "block_bbox": b["bbox"],
                        "key": "t:" + re.sub(r"\d+", "#", " ".join(text.lower().split())),
                        "chars": sum(len(s["text"].strip()) for s in spans),
                    }
                    # "1.<gap>Item text" in one line: remember the number as a
                    # possible automatic label (decided below).
                    first, second = spans[0], spans[1] if len(spans) > 1 else None
                    if (second is not None and _MARKER_ONLY.fullmatch(first["text"].strip())
                            and second["bbox"][0] - first["bbox"][2] > 0.4 * first["size"]):
                        rest_index = all_spans.index(second)
                        element["label"] = first["text"].strip()
                        element["rest"] = "".join(s["text"] for s in all_spans[rest_index:]).strip()
                        element["rest_size"] = second["size"]
                    elements.append(element)
            pages.append((height, elements))

        total_chars = sum(e["chars"] for _, els in pages for e in els if e["type"] == "line")
        if pages and total_chars < 20:
            raise ValueError(
                f"{name} has no selectable text - it is probably a scanned PDF.\n"
                f"This tool needs a PDF with real text (run it through OCR first)."
            )

        # --- Remove headers and footers --------------------------------
        def in_margin(el, height):
            return el["bbox"][3] <= height * PDF_MARGIN_SHARE or el["bbox"][1] >= height * (1 - PDF_MARGIN_SHARE)

        margin_keys = [{e["key"] for e in els if in_margin(e, h)} for h, els in pages]
        seen_on = Counter(k for keys in margin_keys for k in keys)
        n_pages = len(pages)
        repeated = {k for k, c in seen_on.items() if n_pages >= 2 and c >= max(2, (n_pages + 1) // 2)}

        def is_header_footer(el, height):
            if not in_margin(el, height):
                return False
            if el["key"] in repeated:
                return True
            return el["type"] == "line" and _PAGE_NUMBER.fullmatch(el["key"][2:]) is not None

        pages = [(h, [e for e in els if not is_header_footer(e, h)]) for h, els in pages]

        # --- Body font size (most text is set in it) -------------------
        sizes = Counter()
        for _, els in pages:
            for e in els:
                if e["type"] == "line":
                    sizes[round(e["size"] * 2) / 2] += e["chars"]
        body_size = sizes.most_common(1)[0][0] if sizes else 0

        # --- Automatic numbering ---------------------------------------
        # Word doesn't store automatic list/heading numbers and bullets as
        # text, but they are printed in the PDF -- as a separate piece
        # before the item's text. Drop those so they don't show up as
        # differences (and so renumbered lists don't, either). A plain
        # number like "3" is only treated as a label in front of a title,
        # so numbers in table columns are kept.
        def drop_label(marker, following_size):
            if marker.isdigit():
                return body_size > 0 and following_size >= body_size + PDF_HEADING_EXTRA_PT
            return True

        for _, els in pages:
            for i, e in enumerate(els):
                if e["type"] != "line":
                    continue
                if "label" in e:
                    if drop_label(e["label"], e["rest_size"]):
                        e["text"], e["starts"] = e["rest"], True
                    continue
                nxt = els[i + 1] if i + 1 < len(els) else None
                if (nxt is not None and nxt["type"] == "line" and nxt["block"] == e["block"]
                        and _MARKER_ONLY.fullmatch(e["text"])
                        and abs(nxt["bbox"][1] - e["bbox"][1]) < 0.5 * (e["bbox"][3] - e["bbox"][1])
                        and nxt["bbox"][0] - e["bbox"][2] > 0.4 * e["size"]
                        and drop_label(e["text"], nxt["size"])):
                    e["drop"] = True
                    nxt["starts"] = True
            els[:] = [e for e in els if not e.get("drop")]

        # --- Lines -> paragraphs ----------------------------------------
        needle = (split_before or "").strip().lower()

        def starts_paragraph(prev, line):
            if prev is None or line["block"] != prev["block"] or line.get("starts"):
                return True
            if abs(line["size"] - prev["size"]) > 0.6:
                return True
            px0, py0, px1, py1 = prev["bbox"]
            lx0, ly0, lx1, ly1 = line["bbox"]
            height = max(py1 - py0, 1)
            if abs(ly0 - py0) < 0.5 * height:
                return False  # same visual line (e.g. text after a tab)
            if ly0 - py1 > 0.5 * height:
                return True   # extra space between paragraphs
            if _LIST_START.match(line["text"]):
                return True
            if needle and line["text"].lower().startswith(needle):
                return True
            bx0, _, bx1, _ = prev["block_bbox"]
            short = px1 < bx1 - max(0.12 * (bx1 - bx0), 2 * prev["size"])
            return short and prev["text"].endswith(_TERMINAL)

        paragraphs = []   # {"lines": [...]} or {"image": img}
        image_count = 0
        for page_number, (_, els) in enumerate(pages):
            prev = None
            for e in els:
                if e["type"] == "img":
                    image_count += 1
                    paragraphs.append({"image": _pdf_image(e["block"], page_number, image_count)})
                    prev = None
                    continue
                if (prev is None and not e.get("starts") and paragraphs and "lines" in paragraphs[-1]
                        and paragraphs[-1]["page"] == page_number - 1):
                    # First line on a page: continue the last paragraph of the
                    # previous page if that one clearly stopped mid-sentence.
                    last = paragraphs[-1]["lines"][-1]
                    if (abs(last["size"] - e["size"]) <= 0.6
                            and not last["text"].endswith(_TERMINAL)
                            and e["text"][:1].islower()):
                        paragraphs[-1]["lines"].append(e)
                        paragraphs[-1]["page"] = page_number
                        prev = e
                        continue
                if starts_paragraph(prev, e):
                    paragraphs.append({"lines": [e], "page": page_number})
                else:
                    paragraphs[-1]["lines"].append(e)
                prev = e

        items = []
        for p in paragraphs:
            if "image" in p:
                items.append({"text": "", "heading": False, "images": [p["image"]]})
                continue
            text = ""
            for ln in p["lines"]:
                t = ln["text"]
                if text.endswith("-") and t[:1].islower():
                    text += t           # word broken with a hyphen at line end
                elif text:
                    text += " " + t
                else:
                    text = t
            text = _DOT_LEADER.sub(" ", text).strip()  # table-of-contents dots
            if not text:
                continue
            para_sizes = Counter()
            for ln in p["lines"]:
                para_sizes[ln["size"]] += ln["chars"]
            size = para_sizes.most_common(1)[0][0]
            heading = (body_size > 0 and size >= body_size + PDF_HEADING_EXTRA_PT
                       and len(p["lines"]) <= 3 and len(text) <= 200)
            items.append({"text": text, "heading": heading, "images": []})
        return items
    finally:
        doc.close()


def image_thumbnail(data):
    """A tiny greyscale version of a picture, for 'does it look the same?' checks."""
    if Image is None:
        return None
    try:
        with Image.open(io.BytesIO(data)) as im:
            im = im.convert("RGBA")
            flat = Image.new("RGBA", im.size, "white")
            flat.alpha_composite(im)
            return (im.size[0] / max(im.size[1], 1),
                    flat.convert("L").resize((32, 32), Image.Resampling.BOX).tobytes())
    except Exception:
        return None


def _flat_grey(data):
    """Decode a picture to greyscale, with transparent areas shown as white."""
    with Image.open(io.BytesIO(data)) as im:
        im = im.convert("RGBA")
        flat = Image.new("RGBA", im.size, "white")
        flat.alpha_composite(im)
        return flat.convert("L")


def compare_by_appearance(old_img, new_img, old_thumb, new_thumb):
    """
    Decide whether two pictures that aren't pixel-identical show the same
    figure. Returns:
      "resaved"  same size, only compression noise differs (a changed
                 label, line or box is still caught, even a single character)
      "resized"  one is a smaller copy; it looks the same, but small details
                 (e.g. one changed character in a tiny label) can't be verified
      None       the pictures are different
    """
    if old_thumb is None or new_thumb is None:
        return None
    (ratio_a, pix_a), (ratio_b, pix_b) = old_thumb, new_thumb
    if abs(ratio_a - ratio_b) > 0.03 * max(ratio_a, ratio_b):
        return None
    diffs = [abs(x - y) for x, y in zip(pix_a, pix_b)]
    if max(diffs) > 48 or sum(diffs) / len(diffs) > 3.0:
        return None  # clearly a different picture
    try:
        a, b = _flat_grey(old_img["bytes"]), _flat_grey(new_img["bytes"])
    except Exception:
        return None
    if a.size != b.size:
        return "resized"
    # Same size: any pixel that differs strongly is a real edit, not noise.
    strong = ImageChops.difference(a, b).point(lambda v: 255 if v > 96 else 0)
    return "resaved" if strong.getbbox() is None else None


def match_similar_images(items1, items2):
    """
    For PDF comparisons: a picture in the new file that isn't pixel-identical
    to one in the old file but shows the same figure (re-compressed or scaled
    down when the PDF was made) is given the old picture's fingerprint, so it
    counts as unchanged. It is marked in the report with how it was matched.
    """
    if Image is None:
        return
    old_imgs = [img for it in items1 for img in it["images"]]
    new_imgs = [img for it in items2 for img in it["images"]]
    old_hashes = {img["hash"] for img in old_imgs}
    new_hashes = {img["hash"] for img in new_imgs}

    candidates = {}
    for img in old_imgs:
        if img["hash"] not in new_hashes and img["hash"] not in candidates:
            candidates[img["hash"]] = (img, image_thumbnail(img["bytes"]))

    remap, taken = {}, set()
    for img in new_imgs:
        h = img["hash"]
        if h in old_hashes or h in remap:
            continue
        remap[h] = None
        thumb = image_thumbnail(img["bytes"])
        best = None
        for old_hash, (old_img, old_thumb) in candidates.items():
            if old_hash in taken:
                continue
            how = compare_by_appearance(old_img, img, old_thumb, thumb)
            if how == "resaved":
                best = (old_hash, how)
                break
            if how == "resized" and best is None:
                best = (old_hash, how)
        if best:
            remap[h] = best
            taken.add(best[0])
    old_by_hash = {img["hash"]: img for img in old_imgs}
    for img in new_imgs:
        if remap.get(img["hash"]):
            img["hash"], img["match"] = remap[img["hash"]]
            img["old"] = old_by_hash[img["hash"]]


# ===========================================================================
# Old Word (.doc) files
# ===========================================================================
#
# The old Word 97-2003 format (.doc) is a different kind of file that
# python-docx can't read. Such files are converted to .docx first, with
# Microsoft Word (Windows, needs the pywin32 package) or LibreOffice.

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"   # start of every .doc file


def is_legacy_word(path):
    """
    True for files that must be converted before comparing: old .doc files
    (and .rtf / .dot). A password-protected .docx also looks like this from
    the outside; converting it fails with a clear message.
    """
    with open(path, "rb") as f:
        head = f.read(8)
    if head.startswith(b"%PDF-") or head.startswith(b"PK"):
        return False  # PDF, or a normal .docx (a zip file)
    if head != OLE_MAGIC and is_html(path):
        return False  # a web page, even if it was given a .doc name
    return (head == OLE_MAGIC or head.startswith(b"{\\rtf")
            or path.lower().endswith((".doc", ".dot", ".rtf")))


def find_libreoffice():
    """Path of LibreOffice's soffice program, or None if it isn't installed."""
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    candidates = []
    for env in ("ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(env)
        if base:
            candidates.append(os.path.join(base, "LibreOffice", "program", "soffice.exe"))
    candidates.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    return next((c for c in candidates if os.path.isfile(c)), None)


def _docx_name(path, outdir):
    return os.path.join(os.path.abspath(outdir), os.path.splitext(os.path.basename(path))[0] + ".docx")


def _convert_with_word(path, outdir):
    """Convert with Microsoft Word itself (Windows only, needs pywin32)."""
    import pythoncom
    import win32com.client

    out = _docx_name(path, outdir)
    pythoncom.CoInitialize()  # needed because this runs in a background thread
    try:
        # A separate, hidden Word instance: the user's open Word windows are
        # never touched, and Quit() below only closes this instance.
        word = win32com.client.DispatchEx("Word.Application")
        try:
            word.Visible = False
            word.DisplayAlerts = 0  # no pop-up questions
            doc = word.Documents.Open(
                FileName=os.path.abspath(path), ConfirmConversions=False, ReadOnly=True,
                AddToRecentFiles=False, PasswordDocument="", Visible=False,
            )
            try:
                doc.SaveAs(FileName=out, FileFormat=12)  # 12 = Word document (.docx)
            finally:
                doc.Close(SaveChanges=0)
        finally:
            word.Quit(SaveChanges=0)
    finally:
        pythoncom.CoUninitialize()
    if not os.path.isfile(out):
        raise RuntimeError("Word did not produce a .docx file")
    return out


def _convert_with_libreoffice(soffice, path, outdir):
    """Convert with LibreOffice in the background."""
    out = _docx_name(path, outdir)
    # A private settings folder, so this also works while LibreOffice is open.
    profile = pathlib.Path(outdir, "libreoffice-profile").absolute().as_uri()
    result = subprocess.run(
        [soffice, f"-env:UserInstallation={profile}", "--headless", "--norestore",
         "--convert-to", "docx", "--outdir", os.path.abspath(outdir), os.path.abspath(path)],
        capture_output=True, text=True, timeout=300,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if not os.path.isfile(out):
        details = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(details[:300] or "LibreOffice did not produce a .docx file")
    return out


def convert_to_docx(path, outdir, use_word=None):
    """
    Convert an old .doc (or .rtf) file to .docx in `outdir`.
    Returns (path_of_the_docx, name_of_the_program_used).
    Microsoft Word is tried first on Windows, then LibreOffice.
    """
    os.makedirs(outdir, exist_ok=True)
    if use_word is None:
        use_word = os.name == "nt"
    problems = []
    if use_word:
        try:
            return _convert_with_word(path, outdir), "Microsoft Word"
        except ImportError:
            problems.append("using Microsoft Word needs the pywin32 package (pip install pywin32)")
        except Exception as exc:
            problems.append(f"Microsoft Word could not convert it ({exc})")
    soffice = find_libreoffice()
    if soffice:
        try:
            return _convert_with_libreoffice(soffice, path, outdir), "LibreOffice"
        except Exception as exc:
            problems.append(f"LibreOffice could not convert it ({exc})")
    else:
        problems.append("LibreOffice is not installed")
    raise RuntimeError(
        f"{os.path.basename(path)} is an old-format Word file (.doc), or a password-protected "
        f"Word file, which this tool can't read directly, and it could not be converted.\n\n"
        f"Easiest fix: open it in Word, choose File > Save As > Word Document (.docx), "
        f"and compare the .docx instead.\n\n"
        f"Details: " + "; ".join(problems) + "."
    )


# ===========================================================================
# HTML files (.html / .htm, and Word's single-file web pages .mht / .mhtml)
# ===========================================================================

_HTML_BLOCKS = {
    "p", "div", "li", "td", "th", "caption", "dt", "dd", "blockquote", "pre",
    "section", "article", "header", "footer", "main", "aside", "figure", "figcaption",
    "address", "tr", "table", "ul", "ol", "dl", "body", "form", "fieldset", "legend",
    "center", "hr", "h1", "h2", "h3", "h4", "h5", "h6", "nav", "details", "summary",
}
_HTML_HEADINGS = {"h1", "h2", "h3", "h4", "h5", "h6"}
# Content of these is never part of the document text.
_HTML_SKIP = {"script", "style", "head", "noscript", "template", "title", "svg", "math",
              "object", "iframe", "button", "select", "textarea", "xml"}
_HTML_VOID = {"br", "img", "hr", "meta", "link", "input", "area", "base", "col", "embed",
              "param", "source", "track", "wbr"}
_IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tif", ".tiff",
                     ".emf", ".wmf", ".svg")
_MAX_IMAGE_BYTES = 50 * 1024 * 1024


def is_html(path):
    """True for HTML files and Word's single-file web pages (.mht)."""
    if path.lower().endswith((".html", ".htm", ".xhtml", ".mht", ".mhtml")):
        return True
    with open(path, "rb") as f:
        head = f.read(2048)
    head = head.lstrip(codecs.BOM_UTF8).lstrip().lower()
    if head.startswith((b"<!doctype html", b"<html")):
        return True
    if head.startswith(b"<?xml") and b"<html" in head:
        return True
    return head.startswith(b"mime-version") and b"multipart/related" in head and b"text/html" in head


def _decode_html(raw):
    """Bytes of an HTML file -> text, using its BOM or <meta charset> (Word often uses windows-1252)."""
    if raw.startswith(codecs.BOM_UTF8):
        return raw[3:].decode("utf-8", "replace")
    if raw.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        return raw.decode("utf-16", "replace")
    m = re.search(rb"""charset\s*=\s*["']?([A-Za-z0-9_.:-]+)""", raw[:4096], re.I)
    for encoding in ([m.group(1).decode("ascii")] if m else []) + ["utf-8"]:
        try:
            return raw.decode(encoding)
        except (LookupError, UnicodeDecodeError):
            continue
    return raw.decode("cp1252", "replace")


def _read_mht(raw):
    """
    Word's "Single File Web Page" (.mht) is an e-mail-style bundle: the page
    plus its pictures. Returns (html_text, base_location, {location: bytes}).
    """
    msg = email.message_from_bytes(raw, policy=email.policy.default)
    html_text, base, parts = None, "", {}
    for part in msg.walk():
        if part.is_multipart():
            continue
        location = (part.get("Content-Location") or "").strip()
        if part.get_content_type() == "text/html" and html_text is None:
            payload = part.get_payload(decode=True) or b""
            charset = part.get_content_charset()
            try:
                html_text = payload.decode(charset) if charset else _decode_html(payload)
            except (LookupError, UnicodeDecodeError):
                html_text = _decode_html(payload)
            base = location
        elif location:
            parts[location] = part.get_payload(decode=True) or b""
    if html_text is None:
        raise ValueError("this .mht file contains no web page")
    return html_text, base, parts


def _is_big_font(tag, attrs, style):
    """Text set in 14 pt or larger (a title written as a plain paragraph)."""
    m = re.search(r"font-size:([\d.]+)(pt|px)", style)
    if m:
        size = float(m.group(1)) * (0.75 if m.group(2) == "px" else 1.0)
        return size >= 14
    if tag == "font":
        try:
            return int((attrs.get("size") or "").strip()) >= 5
        except ValueError:
            return False
    return False


def _is_page_furniture(tag, attrs, style, classes):
    """
    Page headers/footers and footnotes as Word and LibreOffice write them
    into HTML. They are left out because they aren't compared in .docx
    files either.
    """
    title = (attrs.get("title") or "").lower()
    ident = (attrs.get("id") or "").lower()
    return ((tag == "div" and title in ("header", "footer"))                       # LibreOffice
            or any(k in style for k in ("mso-element:header", "mso-element:footer",   # Word
                                        "mso-element:footnote", "mso-element:endnote",
                                        "mso-footnote-id", "mso-endnote-id"))
            or ident.startswith(("sdfootnote", "sdendnote"))                         # LibreOffice
            or "sdfootnoteanc" in classes or "sdendnoteanc" in classes)


class _HtmlReader(HTMLParser):
    """Turns HTML into paragraph items: one per block (paragraph, heading, list item, table cell)."""

    def __init__(self, load_image):
        super().__init__(convert_charrefs=True)
        self.load_image = load_image
        self.items = []
        self.text = []
        self.images = []
        self.stack = []          # open tags: (tag, skips, is_heading, is_big)
        self.skip = 0            # > 0 while inside content that isn't shown
        self.heading = 0         # > 0 while inside a title/heading
        self.big = 0             # > 0 while inside large-font text
        self.para_heading = False
        self.chars = self.big_chars = 0

    def flush(self):
        text = " ".join("".join(self.text).split())
        if text or self.images:
            big = self.chars > 0 and self.big_chars >= 0.8 * self.chars and len(text) <= 200
            self.items.append({"text": text, "images": self.images,
                               "heading": bool(text) and (self.para_heading or big)})
        self.text, self.images = [], []
        self.chars = self.big_chars = 0

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        style = (attrs.get("style") or "").replace(" ", "").lower()
        classes = (attrs.get("class") or "").lower()
        hidden = ("display:none" in style or "hidden" in attrs
                  # Word writes automatic list numbers as text inside such spans;
                  # .docx files don't contain them, so leave them out here too.
                  or "mso-list:ignore" in style
                  or _is_page_furniture(tag, attrs, style, classes))
        if tag == "br":
            self.text.append("\n")
            return
        if tag == "img":
            if not self.skip and not hidden:
                img = self.load_image(attrs.get("src") or "")
                if img is not None:
                    self.images.append(img)
            return
        if tag in _HTML_BLOCKS:
            self.flush()
        if tag in _HTML_VOID:
            return
        skips = tag in _HTML_SKIP or hidden
        is_heading = tag in _HTML_HEADINGS or "msotitle" in classes or "msosubtitle" in classes
        is_big = _is_big_font(tag, attrs, style)
        self.stack.append((tag, skips, is_heading, is_big))
        self.skip += skips
        self.heading += is_heading
        self.big += is_big
        if tag in _HTML_BLOCKS or is_heading:
            self.para_heading = self.heading > 0

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        if tag not in _HTML_VOID:
            self.handle_endtag(tag)

    def handle_endtag(self, tag):
        if not any(entry[0] == tag for entry in self.stack):
            return  # stray end tag
        while self.stack:
            open_tag, skips, is_heading, is_big = self.stack.pop()
            self.skip -= skips
            self.heading -= is_heading
            self.big -= is_big
            if open_tag == tag:
                break
        if tag in _HTML_BLOCKS or tag in _HTML_HEADINGS:
            self.flush()
            self.para_heading = self.heading > 0

    def handle_data(self, data):
        if not self.skip:
            self.text.append(data)
            n = len(data.strip())
            self.chars += n
            if self.big:
                self.big_chars += n

    def close(self):
        super().close()
        self.flush()


def extract_items_html(path, warnings=None):
    """
    Read an HTML (or .mht) file into paragraph items like extract_items().
    Pictures are loaded from the files next to the page (e.g. Word's
    "<name>_files" folder), from inside an .mht file, or from data: URIs.
    Pictures on the web are not downloaded.
    """
    warnings = warnings if warnings is not None else []
    name = os.path.basename(path)
    with open(path, "rb") as f:
        raw = f.read()
    folder = os.path.dirname(os.path.abspath(path))

    head = raw[:2048].lstrip().lower()
    if head.startswith(b"mime-version") or path.lower().endswith((".mht", ".mhtml")):
        page, base, parts = _read_mht(raw)
    else:
        page, base, parts = _decode_html(raw), "", {}

    missing, remote = [], 0
    cache = {}

    def load_image(src):
        nonlocal remote
        src = src.strip()
        if not src:
            return None
        if src in cache:
            return cache[src]
        data, filename = None, None
        if src.lower().startswith("data:"):
            header, _, payload = src.partition(",")
            try:
                data = base64.b64decode(payload) if ";base64" in header.lower() else unquote(payload).encode("latin-1")
            except Exception:
                data = None
            ext = {"image/jpeg": "jpg", "image/gif": "gif", "image/bmp": "bmp",
                   "image/svg+xml": "svg"}.get(header[5:].split(";")[0].lower(), "png")
            filename = f"embedded-image-{len(cache) + 1}.{ext}"
        elif parts:
            for key in (src, urljoin(base, src) if base else src):
                if key in parts:
                    data = parts[key]
                    break
            else:  # last resort: match by file name only
                wanted = unquote(src.split("?")[0]).replace("\\", "/").rsplit("/", 1)[-1]
                for location, blob in parts.items():
                    if location.replace("\\", "/").rsplit("/", 1)[-1] == wanted:
                        data = blob
                        break
            filename = unquote(src.split("?")[0]).replace("\\", "/").rsplit("/", 1)[-1]
            if data is None:
                missing.append(src)
        elif re.match(r"^[a-z][a-z0-9+.-]*://", src, re.I) and not src.lower().startswith("file:"):
            remote += 1
        else:
            local = unquote(src.split("#")[0].split("?")[0])
            if local.lower().startswith("file:"):
                local = url2pathname(local[5:].lstrip("/") if os.name == "nt" else local[5:])
            full = os.path.normpath(os.path.join(folder, local.replace("/", os.sep)))
            filename = os.path.basename(full)
            if (full.lower().endswith(_IMAGE_EXTENSIONS) and os.path.isfile(full)
                    and os.path.getsize(full) <= _MAX_IMAGE_BYTES):
                with open(full, "rb") as f:
                    data = f.read()
            else:
                missing.append(src)
        img = None
        if data:
            img = {"filename": filename, "partname": filename, "bytes": data,
                   "hash": image_fingerprint(data)}
        cache[src] = img
        return img

    reader = _HtmlReader(load_image)
    reader.feed(page)
    reader.close()

    if missing:
        warnings.append(
            f"{len(missing)} picture(s) used by {name} could not be found (for example "
            f"{missing[0]}). Keep the folder with the page's pictures next to the HTML file."
        )
    if remote:
        warnings.append(f"{remote} picture(s) in {name} are on the web and were not downloaded or compared.")
    return reader.items


def file_kind(path):
    """'pdf', 'html' or 'word' (a .docx; old .doc files are converted before this)."""
    if is_pdf(path):
        return "pdf"
    if is_html(path):
        return "html"
    return "word"


def load_document(path, split_before="", warnings=None):
    """Read a .docx, .pdf or .html file into paragraph items."""
    if is_pdf(path):
        return extract_items_pdf(path, split_before)
    if is_html(path):
        return extract_items_html(path, warnings)
    return extract_items(path)


def guess_mime(filename):
    """Pick an image MIME type from the file extension, for embedding in HTML."""
    ext = filename.lower().rsplit(".", 1)[-1]
    return {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "gif": "image/gif",
        "bmp": "image/bmp",
        "emf": "image/x-emf",
        "wmf": "image/x-wmf",
    }.get(ext, "application/octet-stream")


def to_data_uri(filename, data):
    """Turn raw image bytes into a data: URI that can be shown directly in HTML."""
    encoded = base64.b64encode(data).decode("ascii")
    return f"data:{guess_mime(filename)};base64,{encoded}"


APPROX_IMAGE_LABELS = {
    "resaved": "Unchanged (re-saved copy, checked pixel by pixel)",
    "resized": "Looks unchanged — smaller copy, please check small details",
}


def image_html(img, status):
    """Render one image as an HTML card, tagged with its status (same/added/removed)."""
    label = {"same": "", "added": "New", "removed": "Removed"}[status]
    extra = ""
    if status == "same" and img.get("match"):
        label = APPROX_IMAGE_LABELS[img["match"]]
        extra = f" img-{img['match']}"
    badge = f'<div class="img-badge">{label}</div>' if label else ""
    if status == "same" and img.get("match") == "resized" and img.get("old"):
        # Show the old figure next to the new one so it can be checked by eye.
        return f"""
      <div class="img-inline img-same img-resized">
        {badge}
        <div class="img-pair">
          <div><div class="img-caption">Old version</div>
            <img src="{to_data_uri(img['old']['filename'], img['old']['bytes'])}"></div>
          <div><div class="img-caption">New version (smaller copy)</div>
            <img src="{to_data_uri(img['filename'], img['bytes'])}"></div>
        </div>
      </div>"""
    return f"""
      <div class="img-inline img-{status}{extra}">
        {badge}
        <img src="{to_data_uri(img['filename'], img['bytes'])}">
      </div>"""


def match_images(old_images, new_images):
    """
    Pair up the images of an 'old' paragraph with those of the matching
    'new' paragraph and return a list of (image, status) in display order.
    Images whose content exists on both sides are "same"; images only on
    the old side are "removed"; images only on the new side are "added".
    Matching is by content, not filename, and each copy is matched once —
    so a picture used twice in the old paragraph and once in the new one
    gives one "same" and one "removed".
    """
    new_left = Counter(img["hash"] for img in new_images)
    old_left = Counter(img["hash"] for img in old_images)

    result = []
    for img in old_images:
        if new_left[img["hash"]] > 0:
            new_left[img["hash"]] -= 1
            result.append((img, "same"))
        else:
            result.append((img, "removed"))
    for img in new_images:
        if old_left[img["hash"]] > 0:
            old_left[img["hash"]] -= 1
        else:
            result.append((img, "added"))
    return result


def word_diff_html(old_text, new_text):
    """Return HTML showing word-level differences between two strings."""
    old_words = old_text.split()
    new_words = new_text.split()
    matcher = difflib.SequenceMatcher(None, old_words, new_words)

    parts = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            parts.append(html.escape(" ".join(old_words[i1:i2])))
        elif tag == "delete":
            removed = html.escape(" ".join(old_words[i1:i2]))
            parts.append(f'<span class="removed">{removed}</span>')
        elif tag == "insert":
            added = html.escape(" ".join(new_words[j1:j2]))
            parts.append(f'<span class="added">{added}</span>')
        elif tag == "replace":
            removed = html.escape(" ".join(old_words[i1:i2]))
            added = html.escape(" ".join(new_words[j1:j2]))
            parts.append(f'<span class="removed">{removed}</span>')
            parts.append(f'<span class="added">{added}</span>')
    return " ".join(parts)


def align_block(old_block, new_block):
    """
    Line up the paragraphs of a changed region. Returns a list of
    (old_item_or_None, new_item_or_None) pairs in document order:
    a full pair is an edited paragraph, (old, None) was removed and
    (None, new) was added.

    Paragraphs are paired by how similar their text is, not by position,
    so a brand-new paragraph inserted next to an edited one doesn't throw
    the comparison off.
    """
    n, m = len(old_block), len(new_block)
    old_words = [it["text"].split() for it in old_block]
    new_words = [it["text"].split() for it in new_block]

    def similarity(i, j, matcher):
        if not old_words[i] and not new_words[j]:
            return 1.0  # two image-only paragraphs
        if not old_words[i] or not new_words[j]:
            return 0.0
        matcher.set_seq1(old_words[i])
        if (matcher.real_quick_ratio() < SIMILARITY_THRESHOLD
                or matcher.quick_ratio() < SIMILARITY_THRESHOLD):
            return 0.0
        ratio = matcher.ratio()
        return ratio if ratio >= SIMILARITY_THRESHOLD else 0.0

    pairs = []

    if n * m > MAX_ALIGN_CELLS:
        # Very large changed region: pair by position, but still only call
        # two paragraphs "changed" if they actually resemble each other.
        matcher = difflib.SequenceMatcher(None, autojunk=False)
        for k in range(max(n, m)):
            if k < n and k < m:
                matcher.set_seq2(new_words[k])
                if similarity(k, k, matcher) > 0:
                    pairs.append((old_block[k], new_block[k]))
                    continue
            if k < n:
                pairs.append((old_block[k], None))
            if k < m:
                pairs.append((None, new_block[k]))
        return pairs

    # Similarity of every old paragraph against every new paragraph.
    sim = [[0.0] * m for _ in range(n)]
    matcher = difflib.SequenceMatcher(None, autojunk=False)
    for j in range(m):
        matcher.set_seq2(new_words[j])
        for i in range(n):
            sim[i][j] = similarity(i, j, matcher)

    # Best order-preserving pairing (maximum total similarity).
    score = [[0.0] * (m + 1) for _ in range(n + 1)]
    for i in range(n - 1, -1, -1):
        for j in range(m - 1, -1, -1):
            best = max(score[i + 1][j], score[i][j + 1])
            if sim[i][j] > 0:
                best = max(best, score[i + 1][j + 1] + sim[i][j])
            score[i][j] = best

    i = j = 0
    while i < n and j < m:
        if sim[i][j] > 0 and score[i][j] == score[i + 1][j + 1] + sim[i][j]:
            pairs.append((old_block[i], new_block[j]))
            i += 1
            j += 1
        elif score[i][j] == score[i + 1][j]:
            pairs.append((old_block[i], None))
            i += 1
        else:
            pairs.append((None, new_block[j]))
            j += 1
    pairs.extend((it, None) for it in old_block[i:])
    pairs.extend((None, it) for it in new_block[j:])
    return pairs


def diff_items(items1, items2):
    """
    Compare two ordered lists of paragraph items.
    Returns (rows, counts, img_counts) where rows is a list of
    (kind, html) with kind one of same / changed / added / removed.
    """
    texts1 = [it["text"] for it in items1]
    texts2 = [it["text"] for it in items2]
    matcher = difflib.SequenceMatcher(None, texts1, texts2, autojunk=False)

    rows = []
    counts = {"same": 0, "changed": 0, "added": 0, "removed": 0}
    img_counts = {"same": 0, "added": 0, "removed": 0}

    def add_single(it, kind):
        """A paragraph that exists on one side only (kind: 'added' or 'removed')."""
        text_html = f'<span class="{kind}">{html.escape(it["text"])}</span>' if it["text"] else ""
        img_html = "".join(image_html(img, kind) for img in it["images"])
        img_counts[kind] += len(it["images"])
        rows.append((kind, text_html + img_html))
        counts[kind] += 1

    def add_pair(old_it, new_it):
        """A paragraph that exists on both sides, possibly edited."""
        matched = match_images(old_it["images"], new_it["images"])
        for _, status in matched:
            img_counts[status] += 1
        img_html = "".join(image_html(img, status) for img, status in matched)
        images_changed = any(status != "same" for _, status in matched)

        if old_it["text"] == new_it["text"]:
            text_html = html.escape(old_it["text"])
            # Same words but a picture was added/removed/swapped: still a change.
            kind = "changed" if images_changed else "same"
        else:
            text_html = word_diff_html(old_it["text"], new_it["text"])
            kind = "changed"
        rows.append((kind, text_html + img_html))
        counts[kind] += 1

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                add_pair(items1[i1 + k], items2[j1 + k])
        elif tag == "delete":
            for it in items1[i1:i2]:
                add_single(it, "removed")
        elif tag == "insert":
            for it in items2[j1:j2]:
                add_single(it, "added")
        elif tag == "replace":
            for old_it, new_it in align_block(items1[i1:i2], items2[j1:j2]):
                if new_it is None:
                    add_single(old_it, "removed")
                elif old_it is None:
                    add_single(new_it, "added")
                else:
                    add_pair(old_it, new_it)

    return rows, counts, img_counts


def _flow_tokens(items):
    """Words and pictures of some paragraphs as one stream: (kind, key, starts_paragraph, image)."""
    tokens = []
    for it in items:
        start = True
        for word in it["text"].split():
            tokens.append(("w", word, start, None))
            start = False
        for img in it["images"]:
            tokens.append(("i", img["hash"], start, img))
            start = False
    return tokens


def diff_items_flow(items1, items2):
    """
    Like diff_items(), but ignores where paragraphs are split. Used whenever
    a PDF is involved: a PDF can break one Word paragraph into two (or join
    two), and that alone shouldn't be reported as a change. Only wording,
    order and pictures count. Rows follow the new file's paragraphs.
    """
    rows = []
    counts = {"same": 0, "changed": 0, "added": 0, "removed": 0}
    img_counts = {"same": 0, "added": 0, "removed": 0}

    def render(row):
        statuses = {status for status, _ in row}
        kind = statuses.pop() if len(statuses) == 1 else "changed"
        parts, words, words_status = [], [], None

        def flush_words():
            if words:
                text = html.escape(" ".join(words))
                parts.append(text if words_status == "same" else f'<span class="{words_status}">{text}</span>')
                words.clear()

        for status, tok in row:
            if tok[0] == "w":
                if status != words_status:
                    flush_words()
                    words_status = status
                words.append(tok[1])
            else:
                flush_words()
                words_status = None
                parts.append(image_html(tok[3], status))
                img_counts[status] += 1
                if status == "same" and tok[3].get("match") == "resized":
                    img_counts["check"] = img_counts.get("check", 0) + 1
        flush_words()
        counts[kind] += 1
        rows.append((kind, " ".join(parts)))

    def emit(pieces):
        current = []
        for status, tok, starts in pieces:
            if starts and current:
                render(current)
                current = []
            current.append((status, tok))
        if current:
            render(current)

    def key(it):
        return it["text"] + "\x00" + "|".join(img["hash"] for img in it["images"])

    matcher = difflib.SequenceMatcher(None, [key(i) for i in items1], [key(i) for i in items2], autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            emit([("same", t, t[2]) for t in _flow_tokens(items2[j1:j2])])
            continue
        old_tokens = _flow_tokens(items1[i1:i2])
        new_tokens = _flow_tokens(items2[j1:j2])
        words = difflib.SequenceMatcher(None, [t[:2] for t in old_tokens],
                                        [t[:2] for t in new_tokens], autojunk=False)
        pieces = []
        for op, a1, a2, b1, b2 in words.get_opcodes():
            if op == "equal":
                pieces += [("same", t, t[2]) for t in new_tokens[b1:b2]]
            else:
                pieces += [("removed", t, t[2]) for t in old_tokens[a1:a2]]
                pieces += [("added", t, t[2]) for t in new_tokens[b1:b2]]
        emit(pieces)

    return rows, counts, img_counts


def rows_html(rows):
    return "\n".join(f'<div class="row {kind}">{content}</div>' for kind, content in rows)


# ===========================================================================
# Requirements: splitting a document into requirements and matching them
# ===========================================================================

def split_requirements(items, prefix):
    """
    Group a document's paragraphs into requirements. A requirement starts at
    a paragraph whose text begins with `prefix` (e.g. "Requirement Number")
    and runs until the next such paragraph.

    Returns a list of sections:
        {"id": "12" or None, "name": "Requirement Number 12",
         "items": [...paragraphs...], "words": [...words used for matching...]}

    Each section has a "kind":
      - "req":      a requirement
      - "title":    a title/heading paragraph (plus any loose text under it,
                    up to the next requirement or title). Titles are kept
                    as items of their own so they are never mistaken for
                    part of the requirement above them.
      - "preamble": any text before the first requirement or title

    The requirement's ID is the first word after the prefix
    ("Requirement Number 12: ..." -> "12", "Requirement Number: REQ-7" ->
    "REQ-7"). If the heading paragraph has nothing after the prefix (common
    when the number sits in the next table cell), the next short paragraph
    is used as the ID.
    """
    prefix = prefix.strip()
    needle = prefix.lower()
    sections = []
    current = None

    for idx, it in enumerate(items):
        text = it["text"]
        if needle and text.lower().startswith(needle):
            rest = text[len(prefix):].strip(" \t\n:.#-–—")
            words = rest.split()
            req_id = words[0].rstrip(".:,;)-–—") if words else ""
            body_words = words[1:]
            if not req_id and idx + 1 < len(items):
                nxt = items[idx + 1]["text"].strip()
                if nxt and len(nxt) <= 40 and not nxt.lower().startswith(needle):
                    req_id = nxt
            current = {
                "id": req_id or None,
                "name": f"{prefix} {req_id}".strip(),
                "items": [it],
                "words": list(body_words),
                "kind": "req",
            }
            sections.append(current)
        elif it.get("heading"):
            current = {
                "id": None,
                "name": text if len(text) <= 120 else text[:117] + "...",
                "items": [it],
                "words": text.split(),
                "kind": "title",
            }
            sections.append(current)
        else:
            if current is None:
                current = {"id": None, "name": "Text before the first requirement",
                           "items": [], "words": [], "kind": "preamble"}
                sections.append(current)
            current["items"].append(it)
            current["words"].extend(text.split())
        # Pictures count towards "is this the same requirement?" too.
        current["words"].extend("image:" + img["hash"] for img in it["images"])

    return sections


def _section_similarity(a, b, matcher=None):
    """How alike two requirements' contents are, from 0.0 to 1.0 (ignoring their numbers)."""
    if a["words"] == b["words"]:
        return 1.0
    if not a["words"] or not b["words"]:
        return 0.0
    matcher = matcher or difflib.SequenceMatcher(None, autojunk=False)
    matcher.set_seqs(a["words"], b["words"])
    return matcher.ratio()


def match_requirements(old_secs, new_secs):
    """
    Decide which old requirement corresponds to which new requirement.
    Returns a list of entries in report order:
        {"old": section_or_None, "new": section_or_None, "renumbered": bool}

    Matching goes in passes, most trustworthy first:
      1. Same number AND similar content  -> same requirement.
      2. Different number but (nearly) identical content -> renumbered.
         (This is what keeps the report readable when a requirement is
         inserted in the middle and everything after it shifts by one.)
      3. Same number, content heavily rewritten -> same requirement.
      4. A title that was renamed but still sits next to the same section.
      5. Whatever is left is new (only in the new file) or removed.
    """
    def keyed(secs):
        seen = Counter()
        out = {}
        for sec in secs:
            if sec["kind"] == "preamble":
                key = ("preamble",)
            elif sec["kind"] == "title":
                base = " ".join(sec["items"][0]["text"].lower().split())
                key = ("title", base, seen[("title", base)])
                seen[("title", base)] += 1
            else:
                base = (sec["id"] or "").lower()
                key = ("req", base, seen[base])  # 3rd part separates duplicate numbers
                seen[base] += 1
            out[key] = sec
        return out

    old_by_key, new_by_key = keyed(old_secs), keyed(new_secs)
    partner_of_new = {}   # id(new section) -> (old section, renumbered?)
    used_old = set()

    def pair(old, new, renumbered):
        partner_of_new[id(new)] = (old, renumbered)
        used_old.add(id(old))

    # Pass 1: same number, similar content.
    for key, new in new_by_key.items():
        old = old_by_key.get(key)
        if old is not None and (key == ("preamble",) or
                                _section_similarity(old, new) >= SIMILARITY_THRESHOLD):
            pair(old, new, False)

    # Pass 2: renumbered (content nearly identical under a different number).
    free_old = [s for s in old_secs if id(s) not in used_old and s["words"]]
    free_new = [s for s in new_secs if id(s) not in partner_of_new and s["words"]]
    if len(free_old) * len(free_new) <= MAX_ALIGN_CELLS:
        matcher = difflib.SequenceMatcher(None, autojunk=False)
        for new in free_new:
            # Titles are short, so a reworded title needs a gentler threshold.
            best = None
            best_score = RENUMBER_THRESHOLD if new["kind"] == "req" else SIMILARITY_THRESHOLD
            matcher.set_seq2(new["words"])
            for old in free_old:
                if id(old) in used_old or old["kind"] != new["kind"]:
                    continue
                matcher.set_seq1(old["words"])
                if matcher.real_quick_ratio() < best_score or matcher.quick_ratio() < best_score:
                    continue
                score = matcher.ratio()
                if score >= best_score:
                    best, best_score = old, score
                    if score == 1.0:
                        break
            if best is not None:
                pair(best, new, new["kind"] == "req")

    # Pass 3: same number, even though the content was heavily rewritten.
    for key, new in new_by_key.items():
        old = old_by_key.get(key)
        if id(new) not in partner_of_new and old is not None and id(old) not in used_old:
            pair(old, new, False)

    # Pass 4: a renamed title. An unmatched title in each file counts as the
    # same title when it sits directly before (or after) the same section.
    old_index = {id(s): i for i, s in enumerate(old_secs)}
    for j, new in enumerate(new_secs):
        if new["kind"] != "title" or id(new) in partner_of_new:
            continue
        for step in (1, -1):
            if not 0 <= j + step < len(new_secs):
                continue
            neighbour_old = partner_of_new.get(id(new_secs[j + step]), (None, False))[0]
            if neighbour_old is None:
                continue
            i = old_index[id(neighbour_old)] - step
            if 0 <= i < len(old_secs):
                candidate = old_secs[i]
                if candidate["kind"] == "title" and id(candidate) not in used_old:
                    pair(candidate, new, False)
                    break

    # Report order: the new document's order, with removed requirements
    # slotted in after whatever preceded them in the old document.
    entries = []
    entry_of_old = {}
    for new in new_secs:
        old, renumbered = partner_of_new.get(id(new), (None, False))
        entry = {"old": old, "new": new, "renumbered": renumbered}
        entries.append(entry)
        if old is not None:
            entry_of_old[id(old)] = entry

    previous_entry = None
    for old in old_secs:
        if id(old) in entry_of_old:
            previous_entry = entry_of_old[id(old)]
            continue
        entry = {"old": old, "new": None, "renumbered": False}
        position = 0 if previous_entry is None else next(
            i for i, e in enumerate(entries) if e is previous_entry) + 1
        entries.insert(position, entry)
        previous_entry = entry

    return entries


STATUS_LABELS = {
    "same": "Unchanged", "changed": "Changed", "added": "New",
    "removed": "Removed", "renumbered": "Renumbered", "check": "Check figure",
}

PAGE_TEMPLATE = string.Template("""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Document comparison</title>
<style>
  body { font-family: Arial, sans-serif; max-width: 900px; margin: 40px auto;
         padding: 0 16px; line-height: 1.6; color: #000; }
  h1 { font-size: 20px; }
  h2 { font-size: 16px; margin: 0 0 6px 0; }
  .summary { background: #f5f5f5; padding: 12px 16px; border-radius: 6px;
             margin-bottom: 24px; font-size: 14px; }
  .row { padding: 8px 10px; border-radius: 4px; margin: 6px 0; }
  .row.same { color: #000; }
  .row.changed { background: #fff9e6; }
  .row.added { background: #eaffea; }
  .row.removed { background: #ffecec; }
  .added { background: #c6f6c6; color: #14532d; padding: 1px 3px; border-radius: 3px; }
  .removed { background: #f8c6c6; color: #7f1d1d; text-decoration: line-through;
             padding: 1px 3px; border-radius: 3px; }
  .legend span { margin-right: 16px; }
  .img-inline { display: inline-block; vertical-align: top; margin: 8px 8px 0 0;
                border: 1px solid #ddd; border-radius: 6px; padding: 6px;
                background: #fff; }
  .img-inline img { max-width: 260px; max-height: 260px; display: block; }
  .img-badge { font-size: 11px; font-weight: bold; margin-bottom: 4px; }
  .img-same { border-color: #ddd; }
  .img-added { border-color: #4caf50; }
  .img-added .img-badge { color: #14532d; }
  .img-removed { border-color: #e53935; opacity: 0.7; }
  .img-removed .img-badge { color: #7f1d1d; }
  .img-resaved { border-color: #3b6fd4; }
  .img-resaved .img-badge { color: #1e3a8a; }
  .img-resized { border: 2px dashed #b45309; }
  .img-resized .img-badge { color: #92400e; }
  .img-pair { display: flex; gap: 12px; flex-wrap: wrap; }
  .img-caption { font-size: 11px; color: #555; margin-bottom: 2px; }

  /* Requirement view */
  .badge { display: inline-block; min-width: 86px; text-align: center; font-size: 12px;
           font-weight: bold; padding: 1px 8px; border-radius: 10px; margin-right: 8px;
           border: 1px solid transparent; }
  .badge-same { background: #eee; color: #444; }
  .badge-changed { background: #ffe9a8; color: #6b4a00; }
  .badge-added { background: #c6f6c6; color: #14532d; }
  .badge-removed { background: #f8c6c6; color: #7f1d1d; }
  .badge-renumbered { background: #d6e4ff; color: #1e3a8a; }
  .badge-check { background: #fff; color: #92400e; border: 1px dashed #b45309; }
  .req { border: 1px solid #ddd; border-left-width: 6px; border-radius: 6px;
         margin: 10px 0; padding: 0 10px; }
  .req > summary { cursor: pointer; padding: 8px 0; font-weight: bold; }
  .req .note { font-weight: normal; color: #555; font-size: 13px; margin-left: 6px; }
  .req-same { border-left-color: #ccc; }
  .req-changed { border-left-color: #e0a800; }
  .req-added { border-left-color: #4caf50; }
  .req-removed { border-left-color: #e53935; }
  .req-renumbered { border-left-color: #3b6fd4; }
  .req-check { border-left-color: #b45309; border-left-style: dashed; }
  .kind-title { background: #f3f4f6; margin-top: 22px; }
  .kind-title > summary { font-size: 17px; }
  .title-tag { font-size: 11px; font-weight: bold; text-transform: uppercase; color: #555;
               border: 1px solid #999; border-radius: 3px; padding: 0 5px; margin-right: 8px;
               vertical-align: middle; }
  .changes-list { margin: 0; padding: 0; list-style: none; columns: 2; }
  .changes-list li { margin: 3px 0; break-inside: avoid; }
  .changes-list a { color: #000; }
  .controls { margin: 10px 0 0 0; }
  .controls button { margin-left: 12px; }
  body.hide-same .req-same { display: none; }
</style>
</head>
<body>
  <h1>Comparing: $name1 vs $name2</h1>
  <div class="summary">
    <div class="legend">
      <span class="added">Added text</span>
      <span class="removed">Removed text</span>
    </div>
    $summary_lines
  </div>
  $body
</body>
</html>""")


def build_report(items1, items2, name1, name2, requirement_prefix=DEFAULT_REQUIREMENT_PREFIX,
                 flow=False, notes=()):
    """
    Compare two documents and return (html_string, summary_counts).

    flow=True ignores paragraph breaks (used when a PDF is involved), and
    `notes` are extra lines shown at the top of the report.

    If paragraphs starting with `requirement_prefix` are found, the report is
    organised requirement by requirement. Otherwise (or if the prefix is
    empty) it falls back to one plain paragraph-by-paragraph comparison.
    """
    prefix = (requirement_prefix or "").strip()
    secs1 = split_requirements(items1, prefix) if prefix else []
    secs2 = split_requirements(items2, prefix) if prefix else []
    req_mode = any(s["kind"] == "req" for s in secs1 + secs2)

    summary = {
        "same": 0, "changed": 0, "added": 0, "removed": 0,
        "img_same": 0, "img_added": 0, "img_removed": 0,
        "req_mode": req_mode,
        "req_same": 0, "req_changed": 0, "req_added": 0,
        "req_removed": 0, "req_renumbered": 0, "req_check": 0, "title_check": 0,
        "title_same": 0, "title_changed": 0, "title_added": 0, "title_removed": 0,
        "img_check": 0,
    }

    def tally(counts, img_counts):
        for k in ("same", "changed", "added", "removed"):
            summary[k] += counts[k]
        for k in ("same", "added", "removed"):
            summary["img_" + k] += img_counts[k]
        summary["img_check"] += img_counts.get("check", 0)

    diff = diff_items_flow if flow else diff_items
    notes_html = "".join(f"<i>{html.escape(n)}</i><br>\n    " for n in notes)

    paragraph_lines = (
        "Paragraphs — unchanged: {same}, changed: {changed}, "
        "added: {added}, removed: {removed}<br>\n    "
        "Images — unchanged: {img_same}, added: {img_added}, removed: {img_removed}"
        "{img_check_note}"
    )

    def finish_summary():
        summary_text = dict(summary)
        summary_text["img_check_note"] = (
            f" — <b>{summary['img_check']} unchanged-looking figure(s) are smaller copies; "
            f"please check their details</b> (marked with a dashed border)"
            if summary["img_check"] else "")
        return summary_text

    if not req_mode:
        rows, counts, img_counts = diff(items1, items2)
        tally(counts, img_counts)
        note = ""
        if prefix:
            note = (f"No paragraphs starting with “{html.escape(prefix)}” were found, "
                    f"so this is a plain paragraph-by-paragraph comparison.<br>\n    ")
        page = PAGE_TEMPLATE.substitute(
            name1=html.escape(name1), name2=html.escape(name2),
            summary_lines=notes_html + note + paragraph_lines.format(**finish_summary()),
            body=rows_html(rows),
        )
        return page, summary

    cards = []
    index = []
    for n, entry in enumerate(match_requirements(secs1, secs2)):
        old, new = entry["old"], entry["new"]
        rows, counts, img_counts = diff(old["items"] if old else [],
                                        new["items"] if new else [])
        tally(counts, img_counts)
        edited_rows = sum(1 for kind, _ in rows if kind != "same")

        note = ""
        if old is None:
            status = "added"
        elif new is None:
            status = "removed"
        elif entry["renumbered"]:
            status = "renumbered"
            # The heading line always differs for a renumbered requirement;
            # anything beyond that is a real content change.
            extra = " — content also changed" if edited_rows > 1 else " — content unchanged"
            note = f"was {old['name']}{extra}"
        else:
            status = "same" if edited_rows == 0 else "changed"
            if status == "same" and img_counts.get("check"):
                # Text identical, but a figure could only be compared as a
                # smaller copy: list it so someone looks at it.
                status = "check"

        sec = new or old
        if sec["kind"] == "req":
            summary["req_" + status] += 1
        elif sec["kind"] == "title":
            summary["title_" + status] += 1

        name = html.escape(sec["name"])
        if sec["kind"] == "title":
            name = f'<span class="title-tag">Title</span>{name}'
        note_html = f'<span class="note">({html.escape(note)})</span>' if note else ""
        badge = f'<span class="badge badge-{status}">{STATUS_LABELS[status]}</span>'
        open_attr = "" if status == "same" else " open"
        cards.append(
            f'<details class="req req-{status} kind-{sec["kind"]}" id="req-{n}"{open_attr}>\n'
            f'<summary>{badge}{name}{note_html}</summary>\n'
            f'{rows_html(rows)}\n</details>'
        )
        if status != "same":
            index.append(f'<li><a href="#req-{n}">{badge}{name}</a></li>')

    if index:
        index_html = ('<div class="summary"><h2>What changed</h2>\n'
                      '<ul class="changes-list">\n' + "\n".join(index) + '\n</ul></div>')
    else:
        index_html = '<div class="summary"><h2>What changed</h2>Nothing — all requirements are identical.</div>'

    controls = (
        '<div class="controls">'
        '<label><input type="checkbox" '
        "onchange=\"document.body.classList.toggle('hide-same', this.checked)\"> "
        'Show only requirements that changed</label>'
        "<button onclick=\"document.querySelectorAll('details.req').forEach(function(d){d.open=true})\">Expand all</button>"
        "<button onclick=\"document.querySelectorAll('details.req-same').forEach(function(d){d.open=false})\">Collapse unchanged</button>"
        '</div>'
    )
    requirement_line = (
        "Requirements — unchanged: {req_same}, changed: {req_changed}, "
        "new: {req_added}, removed: {req_removed}, renumbered: {req_renumbered}<br>\n    "
    )
    if summary["req_check"]:
        requirement_line = requirement_line.replace(
            "<br>", ", <b>figure to check: {req_check}</b><br>", 1)
    if any(summary["title_" + k] for k in ("same", "changed", "added", "removed", "check")):
        requirement_line += (
            "Titles — unchanged: {title_same}, changed: {title_changed}, "
            "new: {title_added}, removed: {title_removed}<br>\n    "
        )
    page = PAGE_TEMPLATE.substitute(
        name1=html.escape(name1), name2=html.escape(name2),
        summary_lines=notes_html + (requirement_line + paragraph_lines).format(**finish_summary()) + controls,
        body=index_html + "\n" + "\n".join(cards),
    )
    return page, summary


def run_comparison(file1, file2, output_path, requirement_prefix=DEFAULT_REQUIREMENT_PREFIX):
    """
    Do the full comparison and write the HTML report to output_path.
    Each file may be a Word file (.docx, or old .doc which is converted
    first), a PDF or an HTML page.
    Returns (summary_dict, diagnostic_message_or_None).
    """
    with tempfile.TemporaryDirectory(prefix="compare_docx_") as tmp:
        notes = []
        readable = []
        for tag, path in (("old", file1), ("new", file2)):
            if not is_html(path) and is_legacy_word(path):
                converted, program = convert_to_docx(path, os.path.join(tmp, tag))
                notes.append(f"{os.path.basename(path)} is an old-format Word file (.doc); "
                             f"it was converted to .docx with {program} before comparing.")
                readable.append(converted)
            else:
                readable.append(path)
        return _compare_files(readable[0], readable[1], output_path, requirement_prefix,
                              os.path.basename(file1), os.path.basename(file2), notes)


def _compare_files(file1, file2, output_path, requirement_prefix, name1, name2, notes):
    """run_comparison() for files that are .docx, PDF or HTML."""
    notes = list(notes)
    kind1, kind2 = file_kind(file1), file_kind(file2)
    pdf1, pdf2 = kind1 == "pdf", kind2 == "pdf"
    items1 = load_document(file1, requirement_prefix, notes)
    items2 = load_document(file2, requirement_prefix, notes)
    # Different formats lay paragraphs out differently: compare the words
    # and figures in order, not paragraph by paragraph.
    mixed = pdf1 or pdf2 or kind1 != kind2

    if pdf1 or pdf2:
        notes.append(
            "A PDF is involved, so where paragraphs and lines break is ignored: only the "
            "wording, its order and the figures are compared. Page headers, footers, "
            "page numbers and automatic list/heading numbers are left out (Word doesn't "
            "store automatic numbers as text either). A figure that was re-saved when the PDF "
            "was made counts as unchanged if it shows the same picture; such figures are marked."
        )
    elif mixed:
        notes.append(
            "The files are in different formats, so where paragraphs and lines break is "
            "ignored: only the wording, its order and the figures are compared. A figure "
            "that was re-saved in the other format counts as unchanged if it shows the same "
            "picture; such figures are marked."
        )
    if kind1 != kind2:
        # Vector drawings (EMF/WMF) can't be compared across formats (a PDF
        # draws them as lines, a web page uses a picture copy instead);
        # leave them out instead of calling them removed.
        skipped = 0
        for items in (items1, items2):
            for it in items:
                keep = [img for img in it["images"] if image_thumbnail(img["bytes"]) is not None]
                skipped += len(it["images"]) - len(keep)
                it["images"] = keep
            items[:] = [it for it in items if it["text"] or it["images"]]
        if skipped:
            notes.append(
                f"{skipped} figure(s) are vector drawings (EMF/WMF) that can't be compared "
                f"across file formats; they were left out."
            )
    if mixed:
        match_similar_images(items1, items2)

    page, summary = build_report(
        items1, items2, name1, name2,
        requirement_prefix, flow=mixed, notes=notes,
    )

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(page)

    # Diagnostic: warn if some image files inside a .docx never showed up
    # in the report. (A picture used several times is stored only once, so
    # we compare which files were used, not how many times.)
    def coverage(path, items, kind):
        if kind != "word":
            return 0, 0
        on_disk = list_media_files(path)
        used = {img["partname"].lstrip("/") for it in items for img in it["images"]}
        return len(on_disk), len(on_disk & used)

    on_disk1, placed1 = coverage(file1, items1, kind1)
    on_disk2, placed2 = coverage(file2, items2, kind2)

    diagnostic = None
    if (placed1 < on_disk1 or placed2 < on_disk2) and kind1 == kind2 == "word":
        diagnostic = (
            f"Note: some images inside the files were not matched to any text.\n"
            f"  {name1}: {on_disk1} image file(s) found, {placed1} matched\n"
            f"  {name2}: {on_disk2} image file(s) found, {placed2} matched\n"
            f"This usually means those images are in a header/footer\n"
            f"or a format this tool doesn't recognize yet."
        )
    elif placed1 < on_disk1 or placed2 < on_disk2:
        word_name, on_disk, placed = (name1, on_disk1, placed1) if kind1 == "word" else (name2, on_disk2, placed2)
        diagnostic = (
            f"Note: {on_disk - placed} image file(s) in {word_name} were not compared\n"
            f"(header/footer pictures, vector drawings, or a format this tool doesn't recognize)."
        )

    return summary, diagnostic


# ===========================================================================
# GUI
# ===========================================================================

DOCUMENT_TYPES = [("Word, PDF or HTML", "*.docx *.doc *.pdf *.html *.htm *.mht *.mhtml"),
                  ("Word documents", "*.docx *.doc"), ("Web pages", "*.html *.htm *.mht *.mhtml"),
                  ("PDF files", "*.pdf")]


class CompareApp:
    def __init__(self, root):
        self.root = root
        root.title("Compare Word Documents")
        root.resizable(False, False)

        padding = {"padx": 10, "pady": 6}

        frame = ttk.Frame(root, padding=15)
        frame.grid(row=0, column=0, sticky="nsew")

        # --- Old file ---
        ttk.Label(frame, text="Old / original file (.docx, .doc, .pdf or .html):").grid(row=0, column=0, sticky="w", **padding)
        self.file1_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.file1_var, width=55).grid(row=1, column=0, columnspan=2, sticky="we", padx=10)
        ttk.Button(frame, text="Browse...", command=self.pick_file1).grid(row=1, column=2, padx=10)

        # --- New file ---
        ttk.Label(frame, text="New / updated file (.docx, .doc, .pdf or .html):").grid(row=2, column=0, sticky="w", **padding)
        self.file2_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.file2_var, width=55).grid(row=3, column=0, columnspan=2, sticky="we", padx=10)
        ttk.Button(frame, text="Browse...", command=self.pick_file2).grid(row=3, column=2, padx=10)

        # --- Output file ---
        ttk.Label(frame, text="Save report as:").grid(row=4, column=0, sticky="w", **padding)
        self.output_var = tk.StringVar(value="")
        ttk.Entry(frame, textvariable=self.output_var, width=55).grid(row=5, column=0, columnspan=2, sticky="we", padx=10)
        ttk.Button(frame, text="Browse...", command=self.pick_output).grid(row=5, column=2, padx=10)

        # --- Requirement marker ---
        ttk.Label(frame, text="Each requirement starts with (leave empty for a plain comparison):").grid(
            row=6, column=0, columnspan=3, sticky="w", **padding)
        self.prefix_var = tk.StringVar(value=DEFAULT_REQUIREMENT_PREFIX)
        ttk.Entry(frame, textvariable=self.prefix_var, width=55).grid(row=7, column=0, columnspan=2, sticky="we", padx=10)

        # --- Compare button ---
        self.compare_btn = ttk.Button(frame, text="Compare", command=self.start_compare)
        self.compare_btn.grid(row=8, column=0, columnspan=3, pady=(16, 6))

        # --- Status label ---
        self.status_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.status_var, foreground="#000", wraplength=460, justify="left").grid(
            row=9, column=0, columnspan=3, sticky="w", padx=10
        )

        # --- Open report button (hidden until a report exists) ---
        self.open_btn = ttk.Button(frame, text="Open report", command=self.open_report)
        self.open_btn.grid(row=10, column=0, columnspan=3, pady=(6, 0))
        self.open_btn.grid_remove()
        self.last_output_path = None

    def pick_file1(self):
        path = filedialog.askopenfilename(title="Choose the old file", filetypes=DOCUMENT_TYPES)
        if path:
            self.file1_var.set(path)

    def pick_file2(self):
        path = filedialog.askopenfilename(title="Choose the new file", filetypes=DOCUMENT_TYPES)
        if path:
            self.file2_var.set(path)

    def pick_output(self):
        path = filedialog.asksaveasfilename(
            title="Save report as",
            defaultextension=".html",
            filetypes=[("HTML file", "*.html")],
            initialfile=os.path.basename(self.output_var.get()),
        )
        if path:
            self.output_var.set(path)

    def start_compare(self):
        file1 = self.file1_var.get().strip()
        file2 = self.file2_var.get().strip()
        output = self.output_var.get().strip()

        if not file1 or not file2:
            messagebox.showerror("Missing file", "Please choose both the old and the new file (.docx, .doc, .pdf or .html).")
            return
        if not os.path.isfile(file1):
            messagebox.showerror("File not found", f"Can't find this file:\n{file1}")
            return
        if not os.path.isfile(file2):
            messagebox.showerror("File not found", f"Can't find this file:\n{file2}")
            return
        if pymupdf is None and (is_pdf(file1) or is_pdf(file2)):
            messagebox.showerror(
                "PDF support not installed",
                "Comparing PDF files needs the PyMuPDF package.\n\n"
                "Install it with:  pip install pymupdf")
            return
        if not output:
            messagebox.showerror("Missing output name", "Please enter a name for the output HTML file.")
            return
        if not output.lower().endswith(".html"):
            output += ".html"
            self.output_var.set(output)

        self.open_btn.grid_remove()
        self.compare_btn.config(state="disabled")
        if any(not is_html(f) and is_legacy_word(f) for f in (file1, file2)):
            self.status_var.set("Converting the .doc file(s) and comparing... this can take up to a minute.")
        else:
            self.status_var.set("Comparing... this may take a few seconds for large files.")

        # Run the comparison in a background thread so the window doesn't freeze.
        threading.Thread(target=self._compare_worker, args=(file1, file2, output, self.prefix_var.get().strip()), daemon=True).start()

    def _compare_worker(self, file1, file2, output, prefix):
        try:
            summary, diagnostic = run_comparison(file1, file2, output, prefix)
            self.root.after(0, self._compare_done, summary, diagnostic, output, None)
        except Exception as exc:
            self.root.after(0, self._compare_done, None, None, output, exc)

    def _compare_done(self, summary, diagnostic, output, error):
        self.compare_btn.config(state="normal")

        if error is not None:
            self.status_var.set(f"Something went wrong: {error}")
            messagebox.showerror("Comparison failed", str(error))
            return

        lines = [
            f"Done. Report saved to: {output}",
            f"Paragraphs — unchanged: {summary['same']}, changed: {summary['changed']}, "
            f"added: {summary['added']}, removed: {summary['removed']}",
            f"Images — unchanged: {summary['img_same']}, added: {summary['img_added']}, "
            f"removed: {summary['img_removed']}",
        ]
        if summary["req_mode"]:
            lines.insert(1, (
                f"Requirements — unchanged: {summary['req_same']}, changed: {summary['req_changed']}, "
                f"new: {summary['req_added']}, removed: {summary['req_removed']}, "
                f"renumbered: {summary['req_renumbered']}"))
        elif self.prefix_var.get().strip():
            lines.insert(1, (
                f"No paragraphs starting with \"{self.prefix_var.get().strip()}\" were found, "
                f"so a plain comparison was made."))
        if summary.get("img_check"):
            lines.append(
                f"{summary['img_check']} figure(s) look unchanged but are smaller copies in the PDF - "
                f"please check them in the report (dashed border).")
        if diagnostic:
            lines.append("")
            lines.append(diagnostic)

        self.status_var.set("\n".join(lines))
        self.last_output_path = os.path.abspath(output)
        self.open_btn.grid()

    def open_report(self):
        if self.last_output_path and os.path.isfile(self.last_output_path):
            webbrowser.open(pathlib.Path(self.last_output_path).as_uri())


def main():
    root = tk.Tk()
    CompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()