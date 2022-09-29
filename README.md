## Parse all contents of a docx file with `python-docx`

### Installation
```bash
python3 -m pip install docx-parser
```

### Features:
- `paragraph`: text paragraph, with style_id
- `multipart`: paragraph with image or hyperlink
- `table`: table data with merged_cells

### Examples
- CMD
```bash
docx_parser --help

# parse image as file
docx_parser tests/demo.docx -D tests/media -o tests/out.file.jl

# parse image as base64 string
  docx_parser tests/demo.docx -A base64 -o tests/out.base64.jl
```
- Python
```python
from docx_parser import DocumentParser

infile = 'tests/demo.docx'
doc = DocumentParser(infile)
for _type, item in doc.parse():
    print(_type, item)
```
---

### ToDo
- parse text style: color, bgcolor, font, bold, italic ...
- parse paragraph format
