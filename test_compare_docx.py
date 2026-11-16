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
    for status, summary in re.findall(r'<details class="req req-(\w+)"[^>]*>\n<summary>(.*?)</summary>', page):
        out.append((status, re.sub(r"<[^>]+>", "|", summary).strip("|")))
    return out


R1 = ["Requirement Number 1", "The system shall log every login attempt."]
R2 = ["Requirement Number 2", "The system shall lock an account after five failed attempts."]
R3 = ["Requirement Number 3", "Passwords shall be at least twelve characters long.", "This applies to all user roles."]


def test_split_requirements_ids():
    items = [{"text": t, "images": []} for t in [
        "Document title",
        "Requirement Number 5: The system shall work",
        "Requirement Number: REQ-007", "Body",
        "requirement number 3.2.1.", "Body",
        "Requirement Number", "42", "Body in next cell",
    ]]
    secs = cd.split_requirements(items, "Requirement Number")
    assert secs[0]["preamble"] and len(secs[0]["items"]) == 1
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
    assert re.search(r'req-same" id="req-1">', page) and re.search(r'req-changed" id="req-2" open>', page)
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
