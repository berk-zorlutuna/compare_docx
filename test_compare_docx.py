import base64
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
    expected = dict(same=3, changed=0, added=0, removed=0, img_same=0, img_added=0, img_removed=0)
    assert {k: s[k] for k in expected} == expected
    assert s["req_mode"] is False
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


def test_resaved_identical_figure_is_unchanged(tmp_path):
    """Same pixels, different file bytes (re-saved by Word / another tool)."""
    def build(**save_kwargs):
        def b(doc):
            buf = io.BytesIO()
            Image.new("RGB", (40, 40), "red").save(buf, **save_kwargs)
            buf.seek(0)
            doc.add_paragraph("Fig").add_run().add_picture(buf, width=Inches(1))
        return b
    a, b = build(format="PNG", compress_level=1), build(format="BMP")
    s, d, _ = compare(tmp_path, a, b)
    assert (s["img_same"], s["img_added"], s["img_removed"]) == (1, 0, 0), s
    assert s["same"] == 1 and d is None


def test_one_pixel_difference_is_still_a_change(tmp_path):
    def build(dot):
        def b(doc):
            im = Image.new("RGB", (40, 40), "red")
            if dot:
                im.putpixel((5, 5), (0, 0, 0))
            buf = io.BytesIO(); im.save(buf, "PNG"); buf.seek(0)
            doc.add_paragraph("Fig").add_run().add_picture(buf, width=Inches(1))
        return b
    s, _, _ = compare(tmp_path, build(False), build(True))
    assert (s["img_added"], s["img_removed"]) == (1, 1), s


def test_without_pillow_falls_back_to_bytes(tmp_path, monkeypatch):
    monkeypatch.setattr(cd, "Image", None)
    s, _, _ = compare(tmp_path, with_img("Fig", "red"), with_img("Fig", "red"))
    assert s["img_same"] == 1
    assert cd.image_fingerprint(b"not an image") == cd.image_fingerprint(b"not an image")


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


# ---------- requirements ----------

def req_counts(s):
    return (s["req_same"], s["req_changed"], s["req_added"], s["req_removed"], s["req_renumbered"])


def cards(page):
    """[(status, visible title text)] for each requirement card, in report order."""
    out = []
    for status, summary in re.findall(r'<details class="req req-(\w+)[^"]*"[^>]*>\n<summary>(.*?)</summary>', page):
        out.append((status, re.sub(r"(<[^>]+>)+", "|", summary).strip("|")))
    return out


R1 = ["Requirement Number 1", "The system shall log every login attempt."]
R2 = ["Requirement Number 2", "The system shall lock an account after five failed attempts."]
R3 = ["Requirement Number 3", "Passwords shall be at least twelve characters long.", "This applies to all user roles."]


def test_split_requirements_ids():
    items = [{"text": t, "images": [], "heading": False} for t in [
        "Document title",
        "Requirement Number 5: The system shall work",
        "Requirement Number: REQ-007", "Body",
        "requirement number 3.2.1.", "Body",
        "Requirement Number", "42", "Body in next cell",
    ]]
    secs = cd.split_requirements(items, "Requirement Number")
    assert secs[0]["kind"] == "preamble" and len(secs[0]["items"]) == 1
    assert [s["id"] for s in secs[1:]] == ["5", "REQ-007", "3.2.1", "42"]
    assert secs[1]["words"] == ["The", "system", "shall", "work"]


def test_requirements_unchanged_changed_added_removed(tmp_path):
    new2 = [R2[0], "The system shall lock an account after three failed attempts."]
    r4 = ["Requirement Number 4", "Sessions shall expire after thirty minutes of inactivity."]
    s, _, page = compare(tmp_path, paras("Intro", *R1, *R2, *R3), paras("Intro", *R1, *new2, *r4))
    assert s["req_mode"] is True
    assert req_counts(s) == (1, 1, 1, 1, 0), s
    assert [c[0] for c in cards(page)] == ["same", "same", "changed", "removed", "added"]
    assert '<span class="removed">five</span>' in page and '<span class="added">three</span>' in page
    # unchanged requirements are collapsed, changed ones are open
    assert re.search(r'req-same kind-req" id="req-1">', page) and re.search(r'req-changed kind-req" id="req-2" open>', page)
    # index lists only what changed
    assert page.count('<li><a href="#req-') == 3


def test_requirement_inserted_in_middle_renumbers_rest(tmp_path):
    """Inserting a requirement shifts later numbers; they must not all show as rewritten."""
    new = ["Requirement Number 2", "All data shall be encrypted at rest using approved algorithms."]
    shifted2 = ["Requirement Number 3"] + R2[1:]
    shifted3 = ["Requirement Number 4"] + R3[1:]
    s, _, page = compare(tmp_path, paras(*R1, *R2, *R3), paras(*R1, *new, *shifted2, *shifted3))
    assert req_counts(s) == (1, 0, 1, 0, 2), s
    c = cards(page)
    assert [x[0] for x in c] == ["same", "added", "renumbered", "renumbered"]
    assert "was Requirement Number 2" in c[2][1] and "content unchanged" in c[2][1]


def test_renumbered_and_edited(tmp_path):
    moved = ["Requirement Number 9", "Passwords shall be at least fourteen characters long.", R3[2]]
    s, _, page = compare(tmp_path, paras(*R1, *R3), paras(*R1, *moved))
    assert req_counts(s) == (1, 0, 0, 0, 1), s
    assert "content also changed" in page


def test_short_requirement_renumbered_and_edited(tmp_path):
    old5 = ["Requirement Number 5", "Audit logs shall be retained for one year."]
    old6 = ["Requirement Number 6", "The system shall support single sign-on via SAML."]
    new6 = ["Requirement Number 6", "Audit logs shall be retained for two years."]
    s, _, page = compare(tmp_path, paras(*R1, *old5, *old6), paras(*R1, *new6))
    assert req_counts(s) == (1, 0, 0, 1, 1), s
    assert "was Requirement Number 5" in page


def test_same_number_heavily_rewritten_is_changed(tmp_path):
    rewritten = ["Requirement Number 2", "Backups are taken nightly and stored offsite."]
    s, _, _ = compare(tmp_path, paras(*R1, *R2), paras(*R1, *rewritten))
    assert req_counts(s) == (1, 1, 0, 0, 0), s


def test_requirements_reordered(tmp_path):
    s, _, page = compare(tmp_path, paras(*R1, *R2, *R3), paras(*R3, *R1, *R2))
    assert req_counts(s) == (3, 0, 0, 0, 0), s


def test_requirement_text_in_heading_paragraph(tmp_path):
    s, _, page = compare(
        tmp_path,
        paras("Requirement Number 1: The pump shall start within 2 seconds.",
              "Requirement Number 2: The valve shall close on power loss."),
        paras("Requirement Number 1: The pump shall start within 3 seconds.",
              "Requirement Number 2: The valve shall close on power loss."))
    assert req_counts(s) == (1, 1, 0, 0, 0), s


def test_requirement_number_in_table_cells(tmp_path):
    def build(text):
        def b(doc):
            for num, body in (("1", "Alpha shall beta."), ("2", text)):
                t = doc.add_table(rows=2, cols=2)
                t.cell(0, 0).text = "Requirement Number"; t.cell(0, 1).text = num
                t.cell(1, 0).text = "Description"; t.cell(1, 1).text = body
        return b
    s, _, page = compare(tmp_path, build("Gamma shall delta."), build("Gamma shall epsilon."))
    assert req_counts(s) == (1, 1, 0, 0, 0), s
    assert [c[1] for c in cards(page)] == ["Unchanged|Requirement Number 1", "Changed|Requirement Number 2"]


def test_requirement_image_change(tmp_path):
    def build(color):
        def b(doc):
            paras(*R1)(doc)
            doc.add_paragraph("Requirement Number 2")
            with_img("See figure", color)(doc)
        return b
    s, _, _ = compare(tmp_path, build("red"), build("blue"))
    assert req_counts(s) == (1, 1, 0, 0, 0), s


def test_duplicate_requirement_numbers(tmp_path):
    a = ["Requirement Number 1", "First text about logging events."]
    b = ["Requirement Number 1", "Second text about archiving records."]
    s, _, _ = compare(tmp_path, paras(*a, *b), paras(*a, *b))
    assert req_counts(s) == (2, 0, 0, 0, 0), s


def doc_with_titles(*parts):
    """parts: ("h", text) for a heading, ("t", text) for a Title-style line, else plain text."""
    def build(doc):
        for p in parts:
            if isinstance(p, tuple) and p[0] == "h":
                doc.add_heading(p[1], level=1)
            elif isinstance(p, tuple) and p[0] == "t":
                doc.add_heading(p[1], level=0)
            else:
                doc.add_paragraph(p)
    return build


def test_titles_are_separate_items(tmp_path):
    build = doc_with_titles(("t", "System Requirements"), ("h", "Access control"), *R1, *R2,
                            ("h", "Passwords"), "This chapter covers passwords.", *R3)
    s, _, page = compare(tmp_path, build, build)
    assert req_counts(s) == (3, 0, 0, 0, 0), s
    assert (s["title_same"], s["title_changed"], s["title_added"], s["title_removed"]) == (3, 0, 0, 0)
    assert [c[1] for c in cards(page)] == [
        "Unchanged|Title|System Requirements", "Unchanged|Title|Access control",
        "Unchanged|Requirement Number 1", "Unchanged|Requirement Number 2",
        "Unchanged|Title|Passwords", "Unchanged|Requirement Number 3"]
    # the title is not inside the requirement that precedes it
    secs = cd.split_requirements(cd.extract_items(str(tmp_path / "old.docx")), "Requirement Number")
    assert [len(x["items"]) for x in secs] == [1, 1, 2, 2, 2, 3]


def test_title_change_does_not_touch_neighbouring_requirement(tmp_path):
    old = doc_with_titles(*R1, ("h", "Passwords"), *R3)
    new = doc_with_titles(*R1, ("h", "Password rules"), *R3)
    s, _, page = compare(tmp_path, old, new)
    assert req_counts(s) == (2, 0, 0, 0, 0), s
    assert (s["title_same"], s["title_changed"], s["title_added"], s["title_removed"]) == (0, 1, 0, 0)
    assert '<span class="removed">Passwords</span>' in page
    assert '<span class="added">Password rules</span>' in page


def test_title_added_and_removed(tmp_path):
    old = doc_with_titles(("h", "Legacy interfaces"), *R1, *R2)
    new = doc_with_titles(*R1, ("h", "Account locking"), *R2)
    s, _, page = compare(tmp_path, old, new)
    assert req_counts(s) == (2, 0, 0, 0, 0), s
    assert (s["title_added"], s["title_removed"]) == (1, 1)
    assert [c[0] for c in cards(page)] == ["removed", "same", "added", "same"]


def test_custom_heading_style_and_outline_level(tmp_path):
    def build(doc):
        from docx.enum.style import WD_STYLE_TYPE
        st = doc.styles.add_style("ChapterName", WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = doc.styles["Heading 2"]
        doc.add_paragraph("Chapter one", style="ChapterName")
        doc.add_paragraph("Plain body")
        p = doc.add_paragraph("Outline-level title")
        p._p.get_or_add_pPr().append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="1"/>'))
        q = doc.add_paragraph("Body-level text")
        q._p.get_or_add_pPr().append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="9"/>'))
    items = cd.extract_items(make(tmp_path, "a.docx", build))
    assert [(i["text"], i["heading"]) for i in items] == [
        ("Chapter one", True), ("Plain body", False),
        ("Outline-level title", True), ("Body-level text", False)]


def test_requirement_line_in_heading_style_is_a_requirement(tmp_path):
    build = doc_with_titles(("h", "Requirement Number 1"), "Body one.", ("h", "Requirement Number 2"), "Body two.")
    s, _, _ = compare(tmp_path, build, build)
    assert req_counts(s) == (2, 0, 0, 0, 0) and s["title_same"] == 0


def test_custom_and_empty_prefix(tmp_path):
    f1 = make(tmp_path, "o.docx", paras("REQ-1", "Old text here", "REQ-2", "Stays"))
    f2 = make(tmp_path, "n.docx", paras("REQ-1", "New text here", "REQ-2", "Stays"))
    out = str(tmp_path / "r.html")
    s, _ = cd.run_comparison(f1, f2, out, "REQ-")
    assert req_counts(s) == (1, 1, 0, 0, 0), s
    s, _ = cd.run_comparison(f1, f2, out, "")
    assert s["req_mode"] is False and s["changed"] == 1
    s, _ = cd.run_comparison(f1, f2, out)  # default prefix not present -> plain report
    assert s["req_mode"] is False
    assert "No paragraphs starting with" in open(out, encoding="utf-8").read()


def test_many_requirements_speed(tmp_path):
    import time
    def reqs(n, edit=()):
        out = []
        for i in range(1, n + 1):
            out += [f"Requirement Number {i}",
                    f"The subsystem {i} shall satisfy constraint {i * 7} under condition {i * 13}"
                    + (" revised" if i in edit else "")]
        return out
    t = time.time()
    s, _, _ = compare(tmp_path, paras(*reqs(800)), paras(*reqs(800, edit={5, 400, 799})))
    assert req_counts(s) == (797, 3, 0, 0, 0), s
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


# ---------- PDF ----------

import shutil
import subprocess

pymupdf = pytest.importorskip("pymupdf")


def make_pdf(tmp_path, name, pages, header=None, footer=None, size=(595, 842)):
    """
    pages: list of pages; each page is a list of (text, fontsize) lines or
    ("IMG", png_bytes) entries, laid out top to bottom. (text, fontsize, x)
    places a piece on the previous line at x (same baseline).
    """
    doc = pymupdf.open()
    for n, lines in enumerate(pages, 1):
        page = doc.new_page(width=size[0], height=size[1])
        if header:
            page.insert_text((72, 40), header, fontsize=9)
        if footer:
            page.insert_text((72, size[1] - 30), footer.format(n=n, total=len(pages)), fontsize=9)
        y = 100
        for entry in lines:
            if entry[0] == "IMG":
                page.insert_image(pymupdf.Rect(72, y, 272, y + 100), stream=entry[1])
                y += 115
                continue
            text, fs = entry[0], entry[1]
            if len(entry) == 3:   # same line as previous, at x
                page.insert_text((entry[2], y - fs * 1.3), text, fontsize=fs)
                continue
            if text == "":        # blank line = paragraph gap
                y += 10
                continue
            page.insert_text((72, y), text, fontsize=fs)
            y += fs * 1.3
    path = str(tmp_path / name)
    doc.save(path)
    return path


def texts(items):
    return [("#" if i["heading"] else "") + (i["text"] or "[img]") for i in items]


def test_is_pdf(tmp_path):
    p = make_pdf(tmp_path, "a.pdf", [[("Hello world here", 11)]])
    d = make(tmp_path, "a.docx", paras("x"))
    assert cd.is_pdf(p) and not cd.is_pdf(d)


def test_pdf_paragraphs_headers_footers_and_page_breaks(tmp_path):
    body = 11
    pages = [
        [("Chapter one", 16), ("", 0),
         ("This is the first paragraph of the document.", body), ("", 0),
         ("This paragraph is long and continues on the", body),
         ("next line and then over the page break", body)],
        [("into the second page where it ends.", body), ("", 0),
         ("Requirement Number 1", body), ("", 0),
         ("The unit shall work.", body)],
        [("Last page text here for the check.", body)],
    ]
    p = make_pdf(tmp_path, "a.pdf", pages, header="ACME Spec Rev A", footer="Page {n} of {total}")
    assert texts(cd.extract_items_pdf(p, "Requirement Number")) == [
        "#Chapter one",
        "This is the first paragraph of the document.",
        "This paragraph is long and continues on the next line and then over the page break "
        "into the second page where it ends.",
        "Requirement Number 1",
        "The unit shall work.",
        "Last page text here for the check.",
    ]


def test_pdf_requirement_line_always_starts_a_paragraph(tmp_path):
    p = make_pdf(tmp_path, "a.pdf", [[
        ("The pump shall start within two seconds of the command being received by", 11),
        ("Requirement Number 2", 11),
        ("The valve shall close.", 11)]])
    assert texts(cd.extract_items_pdf(p, "Requirement Number")) == [
        "The pump shall start within two seconds of the command being received by",
        "Requirement Number 2 The valve shall close."]
    # without the marker, the line would have been glued onto the paragraph above
    assert len(cd.extract_items_pdf(p, "")) == 1


def test_pdf_automatic_numbers_dropped_but_typed_numbers_kept(tmp_path):
    p = make_pdf(tmp_path, "a.pdf", [[
        ("3", 16), ("Power supply", 16, 100), ("", 0),          # numbered title: label dropped
        ("1.", 11), ("Over-voltage", 11, 100),                    # list item: label dropped
        ("2.", 11), ("Under-voltage", 11, 100), ("", 0),
        ("Retries", 11), ("3", 11, 200), ("", 0),                 # table-like value: kept
        ("10 seconds is the limit.", 11)]])                       # typed number: kept
    assert texts(cd.extract_items_pdf(p)) == [
        "#Power supply", "Over-voltage", "Under-voltage", "Retries 3", "10 seconds is the limit."]


def test_pdf_dot_leaders_removed(tmp_path):
    p = make_pdf(tmp_path, "a.pdf", [[("Introduction ............................ 3", 11)]])
    assert texts(cd.extract_items_pdf(p)) == ["Introduction 3"]


def test_scanned_pdf_is_rejected(tmp_path):
    p = make_pdf(tmp_path, "scan.pdf", [[("IMG", png("grey", (200, 200)).getvalue())]])
    with pytest.raises(ValueError, match="scanned"):
        cd.extract_items_pdf(p)


def test_password_pdf_is_rejected(tmp_path):
    doc = pymupdf.open()
    doc.new_page().insert_text((72, 72), "Secret text for the test", fontsize=11)
    p = str(tmp_path / "locked.pdf")
    doc.save(p, encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw="pw", owner_pw="pw")
    with pytest.raises(ValueError, match="password"):
        cd.extract_items_pdf(p)


def test_missing_pymupdf_gives_install_hint(tmp_path, monkeypatch):
    p = make_pdf(tmp_path, "a.pdf", [[("Hello world here", 11)]])
    monkeypatch.setattr(cd, "pymupdf", None)
    with pytest.raises(RuntimeError, match="pip install pymupdf"):
        cd.run_comparison(p, p, str(tmp_path / "r.html"))


def rgba_png():
    im = Image.new("RGBA", (120, 60), (0, 0, 0, 0))
    for x in range(20, 100):
        for y in range(10, 50):
            im.putpixel((x, y), (200, 30, 30, 255))
    b = io.BytesIO(); im.save(b, "PNG")
    return b.getvalue()


def test_pdf_transparent_image_matches_word_image(tmp_path):
    data = rgba_png()
    pdf = make_pdf(tmp_path, "a.pdf", [[("See the figure below here.", 11), ("IMG", data)]])
    def build(doc):
        doc.add_paragraph("See the figure below here.")
        doc.add_paragraph().add_run().add_picture(io.BytesIO(data), width=Inches(2))
    docx = make(tmp_path, "a.docx", build)
    s, _ = cd.run_comparison(docx, pdf, str(tmp_path / "r.html"))
    assert (s["same"], s["changed"], s["img_same"], s["img_added"], s["img_removed"]) == (2, 0, 1, 0, 0), s


def figure(label="Pump P-101", size=(600, 300)):
    from PIL import ImageDraw
    im = Image.new("RGB", size, "white")
    d = ImageDraw.Draw(im)
    for i in range(8):
        d.rectangle((20 + i * 70, 80 + (i % 3) * 40, 70 + i * 70, 120 + (i % 3) * 40), outline="black", width=2)
    d.text((30, 30), label, fill="black")
    return im


def encoded(im, fmt="PNG", scale=1.0, **kw):
    if scale != 1.0:
        im = im.resize((int(im.width * scale), int(im.height * scale)), Image.LANCZOS)
    b = io.BytesIO(); im.save(b, fmt, **kw)
    return {"bytes": b.getvalue()}


@pytest.mark.parametrize("new, expected", [
    (encoded(figure(), "JPEG", quality=70), "resaved"),
    (encoded(figure(), "JPEG", scale=0.5, quality=75), "resized"),
    (encoded(figure("Pump P-102")), None),                       # one character changed
    (encoded(figure("Pump P-102"), "JPEG", quality=70), None),   # ...also after re-compression
    (encoded(figure(size=(600, 400))), None),                    # different shape
    (encoded(Image.new("RGB", (600, 300), "navy")), None),       # different picture
])
def test_compare_by_appearance(new, expected):
    old = encoded(figure())
    got = cd.compare_by_appearance(old, new, cd.image_thumbnail(old["bytes"]), cd.image_thumbnail(new["bytes"]))
    assert got == expected


def test_resized_figure_is_flagged_in_report(tmp_path):
    small = encoded(figure(), "JPEG", scale=0.5, quality=75)["bytes"]
    pdf = make_pdf(tmp_path, "a.pdf", [[("Figure text here for test.", 11), ("IMG", small)]])
    def build(doc):
        doc.add_paragraph("Figure text here for test.")
        doc.add_paragraph().add_run().add_picture(io.BytesIO(encoded(figure())["bytes"]), width=Inches(2))
    docx = make(tmp_path, "a.docx", build)
    out = str(tmp_path / "r.html")
    s, _ = cd.run_comparison(docx, pdf, out)
    assert (s["img_same"], s["img_check"], s["img_added"], s["img_removed"]) == (1, 1, 0, 0)
    page = open(out, encoding="utf-8").read()
    assert "please check their details" in page and "img-resized" in page


def test_requirement_with_resized_figure_is_listed_for_checking(tmp_path):
    small = encoded(figure("Pump P-102"), "JPEG", scale=0.5, quality=75)["bytes"]  # label edited!
    pdf = make_pdf(tmp_path, "a.pdf", [[("Requirement Number 1", 11), ("", 0), ("Text of one.", 11), ("", 0),
                                        ("Requirement Number 2", 11), ("", 0), ("Figure below.", 11),
                                        ("IMG", small)]])
    def build(doc):
        paras("Requirement Number 1", "Text of one.", "Requirement Number 2", "Figure below.")(doc)
        doc.add_paragraph().add_run().add_picture(io.BytesIO(encoded(figure())["bytes"]), width=Inches(2))
    docx = make(tmp_path, "a.docx", build)
    out = str(tmp_path / "r.html")
    s, _ = cd.run_comparison(docx, pdf, out)
    assert (s["req_same"], s["req_check"], s["req_changed"]) == (1, 1, 0), s
    assert change_list(out) == [("Check figure", "Requirement Number 2")]
    page = open(out, encoding="utf-8").read()
    assert "figure to check: 1" in page
    assert page.count("<img ") == 2 and "Old version" in page  # old shown next to new


def test_flow_diff_ignores_paragraph_breaks():
    old = [{"text": "Alpha beta gamma delta.", "images": []}, {"text": "Epsilon zeta.", "images": []}]
    new = [{"text": "Alpha beta", "images": []}, {"text": "gamma delta. Epsilon zeta.", "images": []}]
    rows, counts, _ = cd.diff_items_flow(old, new)
    assert counts == {"same": 2, "changed": 0, "added": 0, "removed": 0}
    new[1]["text"] = "gamma DELTA. Epsilon zeta."
    rows, counts, _ = cd.diff_items_flow(old, new)
    assert counts["changed"] == 1 and '<span class="added">DELTA.</span>' in rows[1][1]


# ----- end to end with real PDFs made by LibreOffice from Word files -----

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")


def to_pdf(paths, outdir):
    subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", str(outdir), *map(str, paths)],
                   check=True, capture_output=True, timeout=300)
    return [str(outdir / (os.path.splitext(os.path.basename(p))[0] + ".pdf")) for p in paths]


def spec(version):
    """A requirement document with titles, auto-numbered lists, a table, a figure and header/footer."""
    def build(doc):
        sec = doc.sections[0]
        sec.header.paragraphs[0].text = "ACME Requirement Specification Rev " + "AB"[version - 1]
        sec.footer.paragraphs[0].text = "Company confidential"
        doc.add_heading("System Requirements", 0)
        for ch in range(1, 4):
            title = "Diagnostics" if ch == 2 else f"Chapter {ch}"
            if ch == 2 and version == 2:
                title = "Diagnostics and logging"
            doc.add_heading(title, 1)
            for k in range(1, 5):
                n = (ch - 1) * 4 + k
                if version == 2 and n == 6:
                    continue
                doc.add_paragraph(f"Requirement Number {n}")
                text = ("The unit shall monitor the supply voltage and report any deviation outside the "
                        "permitted band to the supervisory system within the agreed response time. ") * (1 + n % 3)
                if version == 2 and n == 3:
                    text = text.replace("agreed response time", "response time of 200 ms", 1)
                doc.add_paragraph(text.strip())
                if n % 4 == 0:
                    for item in ("Over-voltage", "Under-voltage" if not (version == 2 and n == 8) else "Brown-out"):
                        doc.add_paragraph(item, style="List Number")
                if n == 5:
                    t = doc.add_table(rows=2, cols=2)
                    t.style = "Table Grid"
                    for i, (a, b) in enumerate([("Timeout", "30 s"), ("Retries", "3" if version == 1 else "5")]):
                        t.cell(i, 0).text, t.cell(i, 1).text = a, b
                if n == 7:
                    doc.add_paragraph().add_run().add_picture(
                        io.BytesIO(encoded(figure("Pump P-101" if version == 1 else "Pump P-102"))["bytes"]),
                        width=Inches(3))
        if version == 2:
            doc.add_paragraph("Requirement Number 13")
            doc.add_paragraph("All events shall be logged.")
    return build


@pytest.fixture(scope="module")
def spec_files(tmp_path_factory):
    if not SOFFICE:
        pytest.skip("LibreOffice not installed")
    d = tmp_path_factory.mktemp("spec")
    docx = [make(d, "v1.docx", spec(1)), make(d, "v2.docx", spec(2))]
    return dict(zip(["v1.docx", "v2.docx", "v1.pdf", "v2.pdf"], docx + to_pdf(docx, d)))


def change_list(out):
    page = open(out, encoding="utf-8").read()
    return re.findall(r'<li><a href="#req-\d+"><span class="badge badge-\w+">([^<]+)</span>'
                      r'(?:<span class="title-tag">Title</span>)?([^<]+)</a>', page)


@pytest.mark.parametrize("a, b", [("v1.docx", "v1.pdf"), ("v1.pdf", "v1.docx"), ("v2.docx", "v2.pdf")])
def test_word_vs_its_own_pdf_has_no_changes(spec_files, tmp_path, a, b):
    out = str(tmp_path / "r.html")
    s, _ = cd.run_comparison(spec_files[a], spec_files[b], out)
    assert change_list(out) == []
    assert s["req_changed"] == s["req_added"] == s["req_removed"] == 0
    assert s["img_added"] == s["img_removed"] == 0


@pytest.mark.parametrize("a, b", [("v1.docx", "v2.docx"), ("v1.docx", "v2.pdf"),
                                  ("v1.pdf", "v2.docx"), ("v1.pdf", "v2.pdf")])
def test_all_format_combinations_find_the_same_changes(spec_files, tmp_path, a, b):
    out = str(tmp_path / "r.html")
    s, _ = cd.run_comparison(spec_files[a], spec_files[b], out)
    assert change_list(out) == [
        ("Changed", "Requirement Number 3"),
        ("Changed", "Diagnostics and logging"),
        ("Changed", "Requirement Number 5"),
        ("Removed", "Requirement Number 6"),
        ("Changed", "Requirement Number 7"),
        ("Changed", "Requirement Number 8"),
        ("New", "Requirement Number 13"),
    ]
    assert (s["img_added"], s["img_removed"]) == (1, 1)


# ---------- old Word (.doc) files ----------

import sys
import types


def test_is_legacy_word(tmp_path):
    docx_file = make(tmp_path, "a.docx", paras("x"))
    pdf_file = make_pdf(tmp_path, "a.pdf", [[("Hello world here", 11)]])
    ole = tmp_path / "a.doc"
    ole.write_bytes(cd.OLE_MAGIC + b"\0" * 100)
    rtf = tmp_path / "a.txt"
    rtf.write_bytes(b"{\\rtf1 hello}")
    misnamed = tmp_path / "really_docx.doc"
    misnamed.write_bytes(open(docx_file, "rb").read())
    assert not cd.is_legacy_word(docx_file) and not cd.is_legacy_word(pdf_file)
    assert cd.is_legacy_word(str(ole)) and cd.is_legacy_word(str(rtf))
    assert not cd.is_legacy_word(str(misnamed))  # content wins over the name


def test_no_converter_gives_save_as_advice(tmp_path, monkeypatch):
    doc = tmp_path / "old.doc"
    doc.write_bytes(cd.OLE_MAGIC + b"\0" * 100)
    monkeypatch.setattr(cd, "find_libreoffice", lambda: None)
    with pytest.raises(RuntimeError) as err:
        cd.convert_to_docx(str(doc), str(tmp_path / "out"), use_word=False)
    assert "Save As" in str(err.value) and "LibreOffice is not installed" in str(err.value)


class FakeWord:
    """Stands in for Microsoft Word's automation interface (win32com)."""

    def __init__(self, source_docx, fail_open=False):
        self.calls = []
        self.source_docx = source_docx
        self.fail_open = fail_open
        outer = self

        class Doc:
            def SaveAs(self, FileName, FileFormat):
                outer.calls.append(("SaveAs", FileFormat))
                shutil.copy(outer.source_docx, FileName)

            def Close(self, SaveChanges):
                outer.calls.append(("Close", SaveChanges))

        class Documents:
            def Open(self, **kwargs):
                outer.calls.append(("Open", kwargs))
                if outer.fail_open:
                    raise OSError("The file appears to be corrupted")
                return Doc()

        self.Documents = Documents()

    def Quit(self, SaveChanges):
        self.calls.append(("Quit", SaveChanges))


def install_fake_word(monkeypatch, word):
    com = types.ModuleType("pythoncom")
    com.CoInitialize = lambda: word.calls.append(("CoInitialize",))
    com.CoUninitialize = lambda: word.calls.append(("CoUninitialize",))
    client = types.ModuleType("win32com.client")

    def dispatch_ex(name):
        word.calls.append(("DispatchEx", name))
        return word

    client.DispatchEx = dispatch_ex
    pkg = types.ModuleType("win32com")
    pkg.client = client
    monkeypatch.setitem(sys.modules, "pythoncom", com)
    monkeypatch.setitem(sys.modules, "win32com", pkg)
    monkeypatch.setitem(sys.modules, "win32com.client", client)


def test_word_conversion_uses_own_hidden_instance_read_only(tmp_path, monkeypatch):
    source = make(tmp_path, "src.docx", paras("Requirement Number 1", "Converted text."))
    doc = tmp_path / "old.doc"
    doc.write_bytes(cd.OLE_MAGIC + b"\0" * 100)
    word = FakeWord(source)
    install_fake_word(monkeypatch, word)
    out, program = cd.convert_to_docx(str(doc), str(tmp_path / "out"), use_word=True)
    assert program == "Microsoft Word" and out.endswith("old.docx")
    assert [i["text"] for i in cd.extract_items(out)] == ["Requirement Number 1", "Converted text."]
    names = [c[0] for c in word.calls]
    assert names == ["CoInitialize", "DispatchEx", "Open", "SaveAs", "Close", "Quit", "CoUninitialize"]
    opened = word.calls[2][1]
    assert opened["ReadOnly"] is True and opened["AddToRecentFiles"] is False
    assert opened["PasswordDocument"] == ""        # never shows a password prompt
    assert word.calls[3] == ("SaveAs", 12)          # .docx format
    assert word.calls[4] == ("Close", 0) and word.calls[5] == ("Quit", 0)  # nothing saved back


def test_word_failure_still_closes_word_and_falls_back(tmp_path, monkeypatch):
    source = make(tmp_path, "src.docx", paras("x"))
    doc = tmp_path / "old.doc"
    doc.write_bytes(cd.OLE_MAGIC + b"\0" * 100)
    word = FakeWord(source, fail_open=True)
    install_fake_word(monkeypatch, word)
    monkeypatch.setattr(cd, "find_libreoffice", lambda: None)
    with pytest.raises(RuntimeError) as err:
        cd.convert_to_docx(str(doc), str(tmp_path / "out"), use_word=True)
    assert "Microsoft Word could not convert it (The file appears to be corrupted)" in str(err.value)
    assert [c[0] for c in word.calls][-2:] == ["Quit", "CoUninitialize"]


def test_word_without_pywin32_falls_back_to_libreoffice(tmp_path, monkeypatch):
    doc = tmp_path / "old.doc"
    doc.write_bytes(cd.OLE_MAGIC + b"\0" * 100)
    monkeypatch.setitem(sys.modules, "pythoncom", None)  # import fails
    monkeypatch.setattr(cd, "find_libreoffice", lambda: None)
    with pytest.raises(RuntimeError, match="pywin32"):
        cd.convert_to_docx(str(doc), str(tmp_path / "out"), use_word=True)


@pytest.fixture(scope="module")
def doc_files(spec_files, tmp_path_factory):
    """The same two specs, saved in the old .doc format by LibreOffice."""
    d = tmp_path_factory.mktemp("docs")
    subprocess.run([SOFFICE, "--headless", "--convert-to", "doc", "--outdir", str(d),
                    spec_files["v1.docx"], spec_files["v2.docx"]],
                   check=True, capture_output=True, timeout=300)
    return dict(spec_files, **{"v1.doc": str(d / "v1.doc"), "v2.doc": str(d / "v2.doc")})


def test_doc_vs_its_own_docx_has_no_changes(doc_files, tmp_path):
    out = str(tmp_path / "r.html")
    s, _ = cd.run_comparison(doc_files["v1.doc"], doc_files["v1.docx"], out)
    assert change_list(out) == []
    page = open(out, encoding="utf-8").read()
    assert "Comparing: v1.doc vs v1.docx" in page
    assert "converted to .docx with LibreOffice" in page


@pytest.mark.parametrize("a, b", [("v1.doc", "v2.doc"), ("v1.doc", "v2.docx"),
                                  ("v1.docx", "v2.doc"), ("v1.doc", "v2.pdf"), ("v1.pdf", "v2.doc")])
def test_doc_combinations_find_the_same_changes(doc_files, tmp_path, a, b):
    out = str(tmp_path / "r.html")
    s, _ = cd.run_comparison(doc_files[a], doc_files[b], out)
    assert change_list(out) == [
        ("Changed", "Requirement Number 3"),
        ("Changed", "Diagnostics and logging"),
        ("Changed", "Requirement Number 5"),
        ("Removed", "Requirement Number 6"),
        ("Changed", "Requirement Number 7"),
        ("Changed", "Requirement Number 8"),
        ("New", "Requirement Number 13"),
    ]


# ---------- HTML ----------

def write(tmp_path, name, content, encoding="utf-8"):
    p = tmp_path / name
    p.write_bytes(content.encode(encoding) if isinstance(content, str) else content)
    return str(p)


def html_texts(path, warnings=None):
    return [("#" if i["heading"] else "") + (i["text"] or "[img]") + ("+img" * bool(i["text"] and i["images"]))
            for i in cd.extract_items_html(path, warnings)]


def test_is_html(tmp_path):
    a = write(tmp_path, "a.html", "<p>x</p>")
    b = write(tmp_path, "noext", "\ufeff  <!DOCTYPE html><html><body>x</body></html>")
    c = write(tmp_path, "word.doc", "<html xmlns:o='urn:schemas-microsoft-com:office:office'><body>x</body></html>")
    d = make(tmp_path, "d.docx", paras("x"))
    assert cd.is_html(a) and cd.is_html(b) and cd.is_html(c) and not cd.is_html(d)
    assert not cd.is_legacy_word(c)   # a web page saved as ".doc" is read as HTML, not converted


def test_generic_html_structure(tmp_path):
    p = write(tmp_path, "a.html", """<!DOCTYPE html><html><head><title>Ignored title</title>
      <style>p { color: red }</style><script>var x = "<p>not text</p>";</script></head>
      <body>
        <h1>System &amp; Requirements</h1>
        <div><div>Nested <b>bold</b> text,
           spread over lines.</div></div>
        <p>Line one<br>line two<br/>line&nbsp;three</p>
        <ol><li>First item</li><li>Second <i>item</i></li></ol>
        <table><tr><th>Parameter</th><th>Value</th></tr><tr><td>Timeout</td><td>30 s</td></tr></table>
        <p style="display: none">Hidden text</p>
        <p hidden>Also hidden</p>
        </span></div><p>After stray end tags</p>
        <p class="note">Unclosed paragraph
        <p>Next paragraph</p>
      </body></html>""")
    assert html_texts(p) == [
        "#System & Requirements", "Nested bold text, spread over lines.",
        "Line one line two line three", "First item", "Second item",
        "Parameter", "Value", "Timeout", "30 s", "After stray end tags",
        "Unclosed paragraph", "Next paragraph"]


def test_html_big_font_paragraph_is_a_title(tmp_path):
    p = write(tmp_path, "a.html", """<html><body>
      <p><font size="6" style="font-size: 26pt">Software Requirements</font></p>
      <p style="font-size: 11pt">Body text <span style="font-size:16pt">big word</span> in a sentence.</p>
      <p><span style="font-size: 24px">Big pixel title</span></p></body></html>""")
    assert html_texts(p) == ["#Software Requirements", "Body text big word in a sentence.", "#Big pixel title"]


WORD_HTML = """<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head><meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<style><!-- p.MsoNormal {margin:0cm;} --></style></head>
<body lang=EN-US>
<div class=WordSection1>
<p class=MsoTitle>Software Requirements Specification</p>
<h1 style='mso-list:l0 level1 lfo1'><![if !supportLists]><span style='mso-list:Ignore'>1<span
style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp; </span></span><![endif]>Power supervision</h1>
<p class=MsoNormal>Requirement Number 1</p>
<p class=MsoNormal>The unit shall report the caf\xe9 temperature<a style='mso-footnote-id:ftn1'
href="#_ftn1" name="_ftnref1"><span class=MsoFootnoteReference>[1]</span></a>.</p>
<p class=MsoListParagraph style='mso-list:l1 level1 lfo2'><![if !supportLists]><span
style='mso-list:Ignore'>1.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp; </span></span><![endif]>Over-voltage</p>
<p class=MsoListParagraph style='mso-list:l1 level1 lfo2'><![if !supportLists]><span
style='mso-list:Ignore'>2.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp; </span></span><![endif]>Under-voltage</p>
<p class=MsoNormal><!--[if gte vml 1]><v:shape id="Picture_1" style='width:200pt;height:100pt'>
 <v:imagedata src="spec_files/image001.png" o:title=""/></v:shape><![endif]--><![if !vml]><img
width=267 height=133 src="spec_files/image002.png" v:shapes="Picture_1"><![endif]></p>
</div>
<div style='mso-element:footnote-list'><div style='mso-element:footnote' id=ftn1>
<p class=MsoFootnoteText>[1] A footnote that the .docx comparison doesn't include.</p></div></div>
</body></html>"""


def word_html(tmp_path, picture_color="red"):
    (tmp_path / "spec_files").mkdir(exist_ok=True)
    (tmp_path / "spec_files" / "image002.png").write_bytes(png(picture_color).getvalue())
    return write(tmp_path, "spec.htm", WORD_HTML, "cp1252")


def test_word_html_conventions(tmp_path):
    p = word_html(tmp_path)
    assert html_texts(p) == [
        "#Software Requirements Specification",
        "#Power supervision",                      # automatic heading number left out
        "Requirement Number 1",
        "The unit shall report the café temperature.",   # windows-1252; footnote mark left out
        "Over-voltage", "Under-voltage",           # automatic list numbers left out
        "[img]",                                   # the picture, once (VML copy ignored)
    ]                                              # footnote text left out


def test_word_html_vs_equivalent_docx_has_no_changes(tmp_path):
    p = word_html(tmp_path)
    def build(doc):
        doc.add_heading("Software Requirements Specification", 0)
        doc.add_heading("Power supervision", 1)
        doc.add_paragraph("Requirement Number 1")
        doc.add_paragraph("The unit shall report the café temperature.")
        doc.add_paragraph("Over-voltage", style="List Number")
        doc.add_paragraph("Under-voltage", style="List Number")
        doc.add_paragraph().add_run().add_picture(png("red"), width=Inches(1))
    d = make(tmp_path, "spec.docx", build)
    out = str(tmp_path / "r.html")
    s, _ = cd.run_comparison(d, p, out)
    assert change_list(out) == []
    assert (s["img_same"], s["img_added"], s["img_removed"]) == (1, 0, 0)
    assert s["title_same"] == 2


def test_html_image_change_detected(tmp_path):
    old_dir, new_dir = tmp_path / "old", tmp_path / "new"
    old_dir.mkdir(); new_dir.mkdir()
    s, _ = cd.run_comparison(word_html(old_dir, "red"), word_html(new_dir, "blue"), str(tmp_path / "r.html"))
    assert (s["img_same"], s["img_added"], s["img_removed"]) == (0, 1, 1)


def test_html_image_sources(tmp_path):
    (tmp_path / "pics").mkdir()
    (tmp_path / "pics" / "my figure.png").write_bytes(png("red").getvalue())
    data_uri = "data:image/png;base64," + base64.b64encode(png("blue").getvalue()).decode()
    p = write(tmp_path, "a.html", f"""<html><body>
      <p>Relative <img src="pics/my%20figure.png"></p>
      <p>Embedded <img src="{data_uri}"></p>
      <p>Web <img src="https://example.com/x.png"></p>
      <p>Missing <img src="pics/gone.png"></p>
      <p>Not a picture <img src="../secret.txt"></p></body></html>""")
    warnings = []
    items = cd.extract_items_html(p, warnings)
    assert [len(i["images"]) for i in items] == [1, 1, 0, 0, 0]
    assert any("2 picture(s) used by a.html could not be found" in w for w in warnings), warnings
    assert any("1 picture(s) in a.html are on the web" in w for w in warnings), warnings


def test_mht_single_file_web_page(tmp_path):
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.image import MIMEImage
    msg = MIMEMultipart("related", type="text/html")
    page = MIMEText(WORD_HTML, "html", "windows-1252")
    page["Content-Location"] = "file:///C:/Docs/spec.htm"
    msg.attach(page)
    pic = MIMEImage(png("red").getvalue(), "png")
    pic["Content-Location"] = "file:///C:/Docs/spec_files/image002.png"
    msg.attach(pic)
    p = write(tmp_path, "spec.mht", msg.as_bytes())
    assert cd.is_html(p)
    assert html_texts(p)[-1] == "[img]" and html_texts(p)[3] == "The unit shall report the café temperature."


@pytest.fixture(scope="module")
def html_files(spec_files, tmp_path_factory):
    """The same two specs, saved as web pages by LibreOffice."""
    d = tmp_path_factory.mktemp("html")
    subprocess.run([SOFFICE, "--headless", "--convert-to", "html", "--outdir", str(d),
                    spec_files["v1.docx"], spec_files["v2.docx"]],
                   check=True, capture_output=True, timeout=300)
    return dict(spec_files, **{"v1.html": str(d / "v1.html"), "v2.html": str(d / "v2.html")})


@pytest.mark.parametrize("a, b", [("v1.docx", "v1.html"), ("v1.html", "v1.pdf")])
def test_html_vs_same_document_has_no_changes(html_files, tmp_path, a, b):
    out = str(tmp_path / "r.html")
    s, _ = cd.run_comparison(html_files[a], html_files[b], out)
    assert change_list(out) == []
    assert s["img_added"] == s["img_removed"] == 0


@pytest.mark.parametrize("a, b", [("v1.html", "v2.html"), ("v1.docx", "v2.html"),
                                  ("v1.html", "v2.docx"), ("v1.pdf", "v2.html"), ("v1.html", "v2.pdf")])
def test_html_combinations_find_the_same_changes(html_files, tmp_path, a, b):
    out = str(tmp_path / "r.html")
    cd.run_comparison(html_files[a], html_files[b], out)
    assert change_list(out) == [
        ("Changed", "Requirement Number 3"),
        ("Changed", "Diagnostics and logging"),
        ("Changed", "Requirement Number 5"),
        ("Removed", "Requirement Number 6"),
        ("Changed", "Requirement Number 7"),
        ("Changed", "Requirement Number 8"),
        ("New", "Requirement Number 13"),
    ]
