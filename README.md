# compare_docx

A small point-and-click tool for comparing two Word (`.docx`) files. It writes
an HTML report showing which paragraphs and images were added, removed or
changed, with word-level highlighting.

## Run it

```bash
pip install python-docx
```

```bash
python compare_docx.py
```

Pick the old file, the new file, where to save the report, and click Compare.

## Requirement documents

If the documents contain paragraphs starting with "Requirement Number" (the
text can be changed in the window), the report is organised per requirement:

- each requirement is marked Unchanged, Changed, New, Removed or Renumbered
- a "What changed" list at the top links to every requirement that differs
- unchanged requirements are collapsed, and can be hidden with one tick box
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
