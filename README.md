# compare_docx

A small point-and-click tool for comparing two Word (`.docx`) files. It writes
an HTML report showing which paragraphs and images were added, removed or
changed, with word-level highlighting.

## Run it

```bash
pip install python-docx pillow pymupdf
```

- `python-docx` is always needed.
- `pillow` is optional but recommended: with it, a figure counts as unchanged
  when it looks the same, even if Word re-saved the image file.
- `pymupdf` is only needed for PDF files.

```bash
python compare_docx.py
```

Pick the old file, the new file, where to save the report, and click Compare.

## PDF files

Either file can be a Word file or a PDF, so Word vs Word, PDF vs PDF and
Word vs PDF all work. A PDF only stores lines of text on pages, so when a PDF
is involved the tool:

- rebuilds paragraphs from the line layout, and treats larger-font lines as titles
- ignores where paragraphs and lines break; only wording, order and figures count
- leaves out page headers, footers and page numbers
- leaves out automatic list and heading numbers (Word doesn't store those as
  text either, so renumbering is never reported as a change)
- counts a figure as unchanged when it shows the same picture after being
  re-saved (checked pixel by pixel, so even a one-character label edit is
  caught)
- if the PDF holds a smaller copy of a figure, small details can't be verified
  automatically: the requirement is listed as "Check figure" and the old and new
  figure are shown side by side

Scanned PDFs (pictures of pages, no selectable text) can't be compared; run
them through OCR first. Figures drawn as vector graphics (EMF/WMF drawings,
shapes, SmartArt) can't be compared with a PDF and are left out, with a note.

## Requirement documents

If the documents contain paragraphs starting with "Requirement Number" (the
text can be changed in the window), the report is organised per requirement:

- each requirement is marked Unchanged, Changed, New, Removed or Renumbered
- a "What changed" list at the top links to every requirement that differs
- unchanged requirements are collapsed, and can be hidden with one tick box
- titles (paragraphs in a Word heading or title style) are separate items, so
  a title is never treated as part of the requirement above it
- a requirement whose number changed but whose content is (nearly) the same
  is recognised as renumbered, so inserting one requirement doesn't make
  every later one look rewritten

Leave the box empty, or use documents without such paragraphs, to get a plain
paragraph-by-paragraph comparison.

## Tests

```bash
pip install pytest pillow
```

```bash
pytest
```

The tests import the GUI module, so Python needs tkinter available.
