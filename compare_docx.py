"""
compare_docx.py

A simple point-and-click window for comparing two Word (.docx) files.
Pick the old file, pick the new file, choose where to save the report,
and click Compare. No command line needed.

Text is compared word-by-word, and images are shown inline and matched
by their actual content. The window is built with tkinter, which comes
built into Python; the only extra package needed is python-docx.

Usage:
    python compare_docx.py
"""

import base64
import difflib
import hashlib
import html
import os
import pathlib
import string
import threading
import webbrowser
import zipfile
from collections import Counter

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from docx import Document
from docx.oxml.ns import qn


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


def extract_items(path):
    """
    Walk through a .docx file's raw XML in document order and return a list
    of paragraph items, one per non-empty paragraph:

        {"text": "...", "images": [{"filename": ..., "partname": ...,
                                    "bytes": ..., "hash": ...}, ...]}

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
            images.append({
                "filename": partname.split("/")[-1],
                "partname": partname,
                "bytes": data,
                "hash": hashlib.md5(data).hexdigest(),
            })

        text = "".join(text_parts).strip()
        if not text and not images:
            continue  # skip fully empty paragraphs

        items.append({"text": text, "images": images})

    return items


def list_media_files(path):
    """Return the set of image files physically stored inside the .docx (word/media/)."""
    with zipfile.ZipFile(path) as z:
        return {n for n in z.namelist() if n.startswith("word/media/") and not n.endswith("/")}


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


def image_html(img, status):
    """Render one image as an HTML card, tagged with its status (same/added/removed)."""
    label = {"same": "", "added": "New", "removed": "Removed"}[status]
    badge = f'<div class="img-badge">{label}</div>' if label else ""
    return f"""
      <div class="img-inline img-{status}">
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

    Any text before the first requirement becomes a section with
    "preamble": True.

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
                "preamble": False,
            }
            sections.append(current)
        else:
            if current is None:
                current = {"id": None, "name": "Text before the first requirement",
                           "items": [], "words": [], "preamble": True}
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
      4. Whatever is left is new (only in the new file) or removed.
    """
    def keyed(secs):
        seen = Counter()
        out = {}
        for sec in secs:
            if sec["preamble"]:
                key = ("preamble",)
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
            best, best_score = None, RENUMBER_THRESHOLD
            matcher.set_seq2(new["words"])
            for old in free_old:
                if id(old) in used_old:
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
                pair(best, new, True)

    # Pass 3: same number, even though the content was heavily rewritten.
    for key, new in new_by_key.items():
        old = old_by_key.get(key)
        if id(new) not in partner_of_new and old is not None and id(old) not in used_old:
            pair(old, new, False)

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
    "removed": "Removed", "renumbered": "Renumbered",
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

  /* Requirement view */
  .badge { display: inline-block; min-width: 86px; text-align: center; font-size: 12px;
           font-weight: bold; padding: 1px 8px; border-radius: 10px; margin-right: 8px;
           border: 1px solid transparent; }
  .badge-same { background: #eee; color: #444; }
  .badge-changed { background: #ffe9a8; color: #6b4a00; }
  .badge-added { background: #c6f6c6; color: #14532d; }
  .badge-removed { background: #f8c6c6; color: #7f1d1d; }
  .badge-renumbered { background: #d6e4ff; color: #1e3a8a; }
  .req { border: 1px solid #ddd; border-left-width: 6px; border-radius: 6px;
         margin: 10px 0; padding: 0 10px; }
  .req > summary { cursor: pointer; padding: 8px 0; font-weight: bold; }
  .req .note { font-weight: normal; color: #555; font-size: 13px; margin-left: 6px; }
  .req-same { border-left-color: #ccc; }
  .req-changed { border-left-color: #e0a800; }
  .req-added { border-left-color: #4caf50; }
  .req-removed { border-left-color: #e53935; }
  .req-renumbered { border-left-color: #3b6fd4; }
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


def build_report(items1, items2, name1, name2, requirement_prefix=DEFAULT_REQUIREMENT_PREFIX):
    """
    Compare two documents and return (html_string, summary_counts).

    If paragraphs starting with `requirement_prefix` are found, the report is
    organised requirement by requirement. Otherwise (or if the prefix is
    empty) it falls back to one plain paragraph-by-paragraph comparison.
    """
    prefix = (requirement_prefix or "").strip()
    secs1 = split_requirements(items1, prefix) if prefix else []
    secs2 = split_requirements(items2, prefix) if prefix else []
    req_mode = any(not s["preamble"] for s in secs1 + secs2)

    summary = {
        "same": 0, "changed": 0, "added": 0, "removed": 0,
        "img_same": 0, "img_added": 0, "img_removed": 0,
        "req_mode": req_mode,
        "req_same": 0, "req_changed": 0, "req_added": 0,
        "req_removed": 0, "req_renumbered": 0,
    }

    def tally(counts, img_counts):
        for k in ("same", "changed", "added", "removed"):
            summary[k] += counts[k]
        for k in ("same", "added", "removed"):
            summary["img_" + k] += img_counts[k]

    paragraph_lines = (
        f"Paragraphs — unchanged: {{same}}, changed: {{changed}}, "
        f"added: {{added}}, removed: {{removed}}<br>\n    "
        f"Images — unchanged: {{img_same}}, added: {{img_added}}, removed: {{img_removed}}"
    )

    if not req_mode:
        rows, counts, img_counts = diff_items(items1, items2)
        tally(counts, img_counts)
        note = ""
        if prefix:
            note = (f"No paragraphs starting with “{html.escape(prefix)}” were found, "
                    f"so this is a plain paragraph-by-paragraph comparison.<br>\n    ")
        page = PAGE_TEMPLATE.substitute(
            name1=html.escape(name1), name2=html.escape(name2),
            summary_lines=note + paragraph_lines.format(**summary),
            body=rows_html(rows),
        )
        return page, summary

    cards = []
    index = []
    for n, entry in enumerate(match_requirements(secs1, secs2)):
        old, new = entry["old"], entry["new"]
        rows, counts, img_counts = diff_items(old["items"] if old else [],
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

        sec = new or old
        if not sec["preamble"]:
            summary["req_" + status] += 1

        name = html.escape(sec["name"])
        note_html = f'<span class="note">({html.escape(note)})</span>' if note else ""
        badge = f'<span class="badge badge-{status}">{STATUS_LABELS[status]}</span>'
        open_attr = "" if status == "same" else " open"
        cards.append(
            f'<details class="req req-{status}" id="req-{n}"{open_attr}>\n'
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
    page = PAGE_TEMPLATE.substitute(
        name1=html.escape(name1), name2=html.escape(name2),
        summary_lines=(requirement_line + paragraph_lines).format(**summary) + controls,
        body=index_html + "\n" + "\n".join(cards),
    )
    return page, summary


def run_comparison(file1, file2, output_path, requirement_prefix=DEFAULT_REQUIREMENT_PREFIX):
    """
    Do the full comparison and write the HTML report to output_path.
    Returns (summary_dict, diagnostic_message_or_None).
    """
    items1 = extract_items(file1)
    items2 = extract_items(file2)

    page, summary = build_report(
        items1, items2, os.path.basename(file1), os.path.basename(file2),
        requirement_prefix,
    )

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(page)

    # Diagnostic: warn if some image files inside the .docx never showed up
    # in the report. (A picture used several times is stored only once, so
    # we compare which files were used, not how many times.)
    def coverage(path, items):
        on_disk = list_media_files(path)
        used = {img["partname"].lstrip("/") for it in items for img in it["images"]}
        return len(on_disk), len(on_disk & used)

    on_disk1, placed1 = coverage(file1, items1)
    on_disk2, placed2 = coverage(file2, items2)

    diagnostic = None
    if placed1 < on_disk1 or placed2 < on_disk2:
        diagnostic = (
            f"Note: some images inside the files were not matched to any text.\n"
            f"  {os.path.basename(file1)}: {on_disk1} image file(s) found, {placed1} matched\n"
            f"  {os.path.basename(file2)}: {on_disk2} image file(s) found, {placed2} matched\n"
            f"This usually means those images are in a header/footer\n"
            f"or a format this tool doesn't recognize yet."
        )

    return summary, diagnostic


# ===========================================================================
# GUI
# ===========================================================================

class CompareApp:
    def __init__(self, root):
        self.root = root
        root.title("Compare Word Documents")
        root.resizable(False, False)

        padding = {"padx": 10, "pady": 6}

        frame = ttk.Frame(root, padding=15)
        frame.grid(row=0, column=0, sticky="nsew")

        # --- Old file ---
        ttk.Label(frame, text="Old / original file:").grid(row=0, column=0, sticky="w", **padding)
        self.file1_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.file1_var, width=55).grid(row=1, column=0, columnspan=2, sticky="we", padx=10)
        ttk.Button(frame, text="Browse...", command=self.pick_file1).grid(row=1, column=2, padx=10)

        # --- New file ---
        ttk.Label(frame, text="New / updated file:").grid(row=2, column=0, sticky="w", **padding)
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
        path = filedialog.askopenfilename(title="Choose the old .docx file", filetypes=[("Word documents", "*.docx")])
        if path:
            self.file1_var.set(path)

    def pick_file2(self):
        path = filedialog.askopenfilename(title="Choose the new .docx file", filetypes=[("Word documents", "*.docx")])
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
            messagebox.showerror("Missing file", "Please choose both the old and the new .docx file.")
            return
        if not os.path.isfile(file1):
            messagebox.showerror("File not found", f"Can't find this file:\n{file1}")
            return
        if not os.path.isfile(file2):
            messagebox.showerror("File not found", f"Can't find this file:\n{file2}")
            return
        if not output:
            messagebox.showerror("Missing output name", "Please enter a name for the output HTML file.")
            return
        if not output.lower().endswith(".html"):
            output += ".html"
            self.output_var.set(output)

        self.open_btn.grid_remove()
        self.compare_btn.config(state="disabled")
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