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

## Tests

```bash
pip install pytest pillow
```

```bash
pytest
```

The tests import the GUI module, so Python needs tkinter available.
