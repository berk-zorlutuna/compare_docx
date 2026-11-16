import io
import os
import re

import pytest
from PIL import Image
from docx import Document
from docx.shared import Inches
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

import compare_docx as cd


# ---------- helpers ----------

def png(color, size=(40, 40)):
    buf = io.BytesIO()
    Image.new("RGB", size, color).save(buf, "PNG")
    buf.seek(0)
    return buf


def make(tmp_path, name, build):
    doc = Document()
    build(doc)
    p = str(tmp_path / name)
    doc.save(p)
    return p


def paras(*texts):
    def build(doc):
        for t in texts:
            doc.add_paragraph(t)
    return build


def compare(tmp_path, b1, b2):
    f1 = make(tmp_path, "old.docx", b1)
    f2 = make(tmp_path, "new.docx", b2)
    out = str(tmp_path / "report.html")
    summary, diag = cd.run_comparison(f1, f2, out)
    with open(out, encoding="utf-8") as f:
        page = f.read()
    return summary, diag, page


def rows(page):
    return re.findall(r'<div class="row (\w+)">', page)


# ---------- extraction ----------

def test_extract_text_skips_empty(tmp_path):
    f = make(tmp_path, "a.docx", paras("One", "", "   ", "Two"))
    assert [i["text"] for i in cd.extract_items(f)] == ["One", "Two"]


def test_extract_table_and_nested_table(tmp_path):
    def build(doc):
        doc.add_paragraph("Before")
        t = doc.add_table(rows=1, cols=2)
        t.cell(0, 0).text = "Cell A"
        inner = t.cell(0, 1).add_table(rows=1, cols=1)
        inner.cell(0, 0).text = "Nested"
        inner.cell(0, 0).paragraphs[0].add_run().add_picture(png("red"), width=Inches(1))
        doc.add_paragraph("After")
    items = cd.extract_items(make(tmp_path, "a.docx", build))
    assert [i["text"] for i in items] == ["Before", "Cell A", "Nested", "After"]
    assert len(items[2]["images"]) == 1


def test_extract_image_only_paragraph(tmp_path):
    def build(doc):
        doc.add_picture(png("red"), width=Inches(1))
    items = cd.extract_items(make(tmp_path, "a.docx", build))
    assert len(items) == 1 and items[0]["text"] == "" and len(items[0]["images"]) == 1


def test_extract_legacy_vml_image(tmp_path):
    def build(doc):
        doc.add_picture(png("red"), width=Inches(1))
        # Rewrite the modern drawing into a legacy VML <v:imagedata r:id=...>
        p = doc.paragraphs[-1]
        blip = [e for e in p._p.iter() if e.tag.endswith("}blip")][0]
        rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
        r = p.runs[0]._r
        for child in list(r):
            r.remove(child)
        pict = parse_xml(
            f'<w:pict {nsdecls("w", "r")} xmlns:v="urn:schemas-microsoft-com:vml">'
            f'<v:shape><v:imagedata r:id="{rid}"/></v:shape></w:pict>')
        r.append(pict)
    items = cd.extract_items(make(tmp_path, "a.docx", build))
    assert len(items) == 1 and len(items[0]["images"]) == 1


def test_tabs_and_line_breaks_separate_words(tmp_path):
    def build(doc):
        p = doc.add_paragraph()
        r = p.add_run("Hello")
        r.add_tab()
        r.add_text("World")
        r.add_break()
        r.add_text("Again")
    items = cd.extract_items(make(tmp_path, "a.docx", build))
    assert items[0]["text"].split() == ["Hello", "World", "Again"]


def test_textbox_text_not_duplicated(tmp_path):
    def build(doc):
        p = doc.add_paragraph("Outer")
        txbx = parse_xml(
            f'<w:r {nsdecls("w")}><w:pict><w:txbxContent>'
            f'<w:p><w:r><w:t>Inside box</w:t></w:r></w:p>'
            f'</w:txbxContent></w:pict></w:r>')
        p._p.append(txbx)
    items = cd.extract_items(make(tmp_path, "a.docx", build))
    all_text = " | ".join(i["text"] for i in items)
    assert all_text.count("Inside box") == 1, all_text


# ---------- text comparison ----------

def test_identical(tmp_path):
    s, d, page = compare(tmp_path, paras("A", "B", "C"), paras("A", "B", "C"))
    assert s == dict(same=3, changed=0, added=0, removed=0, img_same=0, img_added=0, img_removed=0)
    assert d is None


def test_added_removed_changed(tmp_path):
    s, _, page = compare(
        tmp_path,
        paras("Intro", "The quick brown fox", "Gone soon", "End"),
        paras("Intro", "The quick red fox", "End", "Brand new"))
    assert (s["same"], s["changed"], s["added"], s["removed"]) == (2, 1, 1, 1)
    assert '<span class="removed">brown</span>' in page
    assert '<span class="added">red</span>' in page


def test_both_empty(tmp_path):
    s, _, page = compare(tmp_path, paras(), paras())
    assert s["same"] == s["changed"] == s["added"] == s["removed"] == 0


def test_html_escaping(tmp_path):
    s, _, page = compare(tmp_path, paras("<script>alert(1)</script> & co"),
                         paras("<script>alert(2)</script> & co"))
    assert "<script>" not in page
    assert "&lt;script&gt;" in page


def test_unicode(tmp_path):
    s, _, page = compare(tmp_path, paras("Çağrı öğle yemeği 日本語"), paras("Çağrı akşam yemeği 日本語"))
    assert s["changed"] == 1 and "akşam" in page and "日本語" in page


def test_insert_before_edited_paragraph_alignment(tmp_path):
    """New paragraph inserted right before an edited one: the edited paragraph
    should still be diffed against its old version, not against the insert."""
    s, _, page = compare(
        tmp_path,
        paras("Intro", "The budget for 2024 is ten thousand dollars in total", "End"),
        paras("Intro", "Completely unrelated new sentence here",
              "The budget for 2025 is ten thousand dollars in total", "End"))
    assert (s["changed"], s["added"], s["removed"]) == (1, 1, 0), s
    assert '<span class="removed">2024</span>' in page
    assert '<span class="added">2025</span>' in page


def test_unrelated_paragraphs_are_removed_plus_added(tmp_path):
    s, _, page = compare(tmp_path, paras("Intro", "Apples and oranges are fruit", "End"),
                         paras("Intro", "Meeting moved to Thursday afternoon", "End"))
    assert (s["changed"], s["added"], s["removed"]) == (0, 1, 1), s
    assert rows(page) == ["same", "removed", "added", "same"]


def test_several_edits_and_inserts_in_one_block(tmp_path):
    s, _, page = compare(
        tmp_path,
        paras("Start",
              "The first clause covers payment terms and deadlines",
              "The second clause covers delivery and shipping costs",
              "This old clause is being dropped entirely now",
              "Finish"),
        paras("Start",
              "A freshly written opening statement goes here",
              "The first clause covers payment terms and due dates",
              "Another unrelated insertion between two clauses",
              "The second clause covers delivery and freight costs",
              "Finish"))
    assert rows(page) == ["same", "added", "changed", "added", "changed", "removed", "same"], rows(page)
    assert '<span class="added">freight</span>' in page


def test_very_large_changed_block_falls_back(tmp_path, monkeypatch):
    monkeypatch.setattr(cd, "MAX_ALIGN_CELLS", 1)
    s, _, _ = compare(tmp_path, paras("Keep", "alpha beta gamma delta", "one two three four", "End"),
                      paras("Keep", "alpha beta gamma epsilon", "totally different words here", "End"))
    assert (s["changed"], s["added"], s["removed"]) == (1, 1, 1), s


def test_tab_stop_definitions_are_not_text(tmp_path):
    def build(doc):
        p = doc.add_paragraph("Plain")
        p.paragraph_format.tab_stops.add_tab_stop(Inches(2))
    items = cd.extract_items(make(tmp_path, "a.docx", build))
    assert items[0]["text"] == "Plain"


def test_modern_textbox_fallback_copy_ignored(tmp_path):
    """Word saves text boxes twice (mc:Choice + mc:Fallback); report them once."""
    def build(doc):
        p = doc.add_paragraph("Outer")
        box = ('<w:txbxContent><w:p><w:r><w:t>Inside box</w:t></w:r></w:p></w:txbxContent>')
        ac = parse_xml(
            f'<w:r {nsdecls("w")} '
            f'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
            f'<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>{box}</w:drawing></mc:Choice>'
            f'<mc:Fallback><w:pict>{box}</w:pict></mc:Fallback></mc:AlternateContent></w:r>')
        p._p.append(ac)
    items = cd.extract_items(make(tmp_path, "a.docx", build))
    assert [i["text"] for i in items] == ["Outer", "Inside box"]


def test_duplicate_paragraphs(tmp_path):
    s, _, _ = compare(tmp_path, paras("x", "x", "x"), paras("x", "x"))
    assert (s["same"], s["removed"]) == (2, 1)


# ---------- images ----------

def with_img(text, *colors):
    def build(doc):
        p = doc.add_paragraph(text)
        for c in colors:
            p.add_run().add_picture(png(c), width=Inches(1))
    return build


def test_image_unchanged(tmp_path):
    s, d, page = compare(tmp_path, with_img("Fig", "red"), with_img("Fig", "red"))
    assert (s["img_same"], s["img_added"], s["img_removed"]) == (1, 0, 0)
    assert d is None
    assert page.count("<img ") == 1


def test_image_replaced(tmp_path):
    s, _, page = compare(tmp_path, with_img("Fig", "red"), with_img("Fig", "blue"))
    assert (s["img_same"], s["img_added"], s["img_removed"]) == (0, 1, 1)
    assert "img-added" in page and "img-removed" in page
    assert rows(page) == ["changed"]


def test_image_replaced_marks_row_as_changed(tmp_path):
    """Text identical but picture swapped: the row shouldn't be reported 'unchanged'."""
    s, _, page = compare(tmp_path, with_img("Fig", "red"), with_img("Fig", "blue"))
    assert s["same"] == 0, s


def test_image_added_and_removed(tmp_path):
    s, _, _ = compare(tmp_path, with_img("Fig"), with_img("Fig", "red"))
    assert (s["img_same"], s["img_added"], s["img_removed"]) == (0, 1, 0)
    s, _, _ = compare(tmp_path, with_img("Fig", "red"), with_img("Fig"))
    assert (s["img_same"], s["img_added"], s["img_removed"]) == (0, 0, 1)


def test_image_only_paragraph_swapped(tmp_path):
    s, _, _ = compare(tmp_path, with_img("", "red"), with_img("", "blue"))
    assert (s["img_added"], s["img_removed"]) == (1, 1)


def test_image_moved_to_other_paragraph(tmp_path):
    def b1(doc):
        with_img("A", "red")(doc); doc.add_paragraph("B")
    def b2(doc):
        doc.add_paragraph("A"); with_img("B", "red")(doc)
    s, _, _ = compare(tmp_path, b1, b2)
    # Reported as removed+added (no 'moved' concept) - documenting behaviour
    assert (s["img_same"], s["img_added"], s["img_removed"]) == (0, 1, 1)


def test_same_image_twice_in_paragraph_then_once(tmp_path):
    s, _, _ = compare(tmp_path, with_img("Fig", "red", "red"), with_img("Fig", "red"))
    assert (s["img_same"], s["img_added"], s["img_removed"]) == (1, 0, 1), s
    s, _, _ = compare(tmp_path, with_img("Fig", "red"), with_img("Fig", "red", "red"))
    assert (s["img_same"], s["img_added"], s["img_removed"]) == (1, 1, 0), s


def test_same_image_reused_no_false_diagnostic(tmp_path):
    """Same picture used in two paragraphs is stored once in word/media;
    that shouldn't trigger the 'images not matched' warning."""
    def build(doc):
        with_img("A", "red")(doc)
        with_img("B", "red")(doc)
    s, d, _ = compare(tmp_path, build, build)
    assert s["img_same"] == 2
    assert d is None, d


def test_header_image_triggers_diagnostic(tmp_path):
    def build(doc):
        doc.add_paragraph("Body")
        doc.sections[0].header.paragraphs[0].add_run().add_picture(png("red"), width=Inches(1))
    s, d, _ = compare(tmp_path, build, build)
    assert d is not None and "1 image file(s) found, 0 matched" in d


def test_large_document_speed(tmp_path):
    import time
    a = [f"Paragraph number {i} with some filler words to compare" for i in range(3000)]
    b = list(a)
    for i in range(0, 3000, 50):
        b[i] = b[i] + " edited"
    t = time.time()
    s, _, _ = compare(tmp_path, paras(*a), paras(*b))
    assert s["changed"] == 60 and s["same"] == 2940
    assert time.time() - t < 30


# ---------- misc ----------

def test_guess_mime():
    assert cd.guess_mime("a.PNG") == "image/png"
    assert cd.guess_mime("a.jpeg") == "image/jpeg"
    assert cd.guess_mime("a.tiff") == "application/octet-stream"
    assert cd.guess_mime("noext") == "application/octet-stream"


def test_non_docx_raises(tmp_path):
    bad = tmp_path / "bad.docx"
    bad.write_text("not a zip")
    good = make(tmp_path, "g.docx", paras("x"))
    with pytest.raises(Exception):
        cd.run_comparison(str(bad), good, str(tmp_path / "o.html"))


def test_report_title_uses_file_names(tmp_path):
    _, _, page = compare(tmp_path, paras("x"), paras("x"))
    title = re.search(r"<h1>(.*?)</h1>", page).group(1)
    assert title == "Comparing: old.docx vs new.docx", title
