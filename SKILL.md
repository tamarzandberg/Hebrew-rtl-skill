---
name: hebrew-rtl
description: >
  Use this skill whenever the user requests a document, presentation, report, or any written output in Hebrew (עברית).
  Trigger on any of the following: the user writes their request in Hebrew, mentions "בעברית", asks for a מסמך, דוח, מצגת, סיכום, מכתב, or any other deliverable and the content is in Hebrew.
  This skill ensures correct RTL (right-to-left) layout, Hebrew-compatible fonts, proper punctuation placement per Hebrew keyboard conventions, and correct use of hyphens and dashes.
  Always use this skill when producing .docx, .pptx, .pdf, or HTML/web content in Hebrew — even if the user does not explicitly mention RTL or formatting. Do NOT skip this skill just because the user's request was brief.
---

# Hebrew RTL Document Skill

---

## Core Principles

### 1. Direction & Alignment
- All Hebrew text: RTL, right-aligned
- Mixed Hebrew-English: Hebrew is primary; English words flow inline via BiDi
- **Tables, info-boxes, timelines, grids**: always start from the RIGHT — first column/item on the right, last on the left
- Lists: bullets on the right side

### 2. Font
- **Arial** — most reliable for docx/pptx
- **Heebo** — preferred for HTML/web
- Always set `w:hint="cs"` (complex script) on font runs in docx

### 3. Punctuation
- Hyphen: `-` (U+002D) — **NEVER** em-dash `—` or en-dash `–`
- Colon `:`, semicolon `;`, comma `,` — standard
- Quotation marks: `"..."` — not curly quotes

### 4. Numbers and Dates
- Numbers stay LTR within RTL text (BiDi automatic)
- Dates: `DD.MM.YYYY` or `15 במרץ 2025`

---

## Word (.docx) — CRITICAL FINDINGS

### The correct RTL pattern (reverse-engineered from a real Hebrew Word document):

```xml
<!-- Paragraph: ONLY <w:bidi/> — NO w:val attribute, NO <w:jc w:val="right"/> -->
<w:p>
  <w:pPr>
    <w:bidi/>
    <w:rPr><w:rtl/></w:rPr>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rFonts w:hint="cs"/>
      <w:rtl/>
    </w:rPr>
    <w:t xml:space="preserve">הטקסט כאן</w:t>
  </w:r>
</w:p>
```

### What does NOT work:
- ❌ `<w:bidi w:val="1"/>` — must be `<w:bidi/>` with no attribute
- ❌ `<w:jc w:val="right"/>` — Word does NOT use this for Hebrew RTL
- ❌ docx-js `bidirectional: true` + `AlignmentType.RIGHT` — insufficient
- ❌ python-docx paragraph alignment — does not properly set RTL

### The only reliable approach: inject XML into a real Hebrew Word template

Use a `.docx` created by Word with Hebrew locale as base (preserves correct `settings.xml` and `styles.xml`), then inject content via Python + zipfile:

```python
import zipfile, os, re

def build_hebrew_docx(paragraphs_xml, base_docx, output_path):
    os.makedirs('/tmp/heb_work', exist_ok=True)
    with zipfile.ZipFile(base_docx, 'r') as z:
        z.extractall('/tmp/heb_work')

    orig = open('/tmp/heb_work/word/document.xml', encoding='utf-8').read()
    ns_part = orig[:orig.index('<w:body>')]
    sectPr = re.search(r'<w:sectPr.*?</w:sectPr>', orig, re.DOTALL).group()

    body = '\n'.join(paragraphs_xml)
    new_doc = f'{ns_part}<w:body>{body}\n{sectPr}</w:body></w:document>'
    open('/tmp/heb_work/word/document.xml', 'w', encoding='utf-8').write(new_doc)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk('/tmp/heb_work'):
            for file in files:
                fp = os.path.join(root, file)
                zout.write(fp, os.path.relpath(fp, '/tmp/heb_work'))
```

### Paragraph builder functions:

```python
def heb_para(text, bold=False, size=24, color=None, space_before=None, space_after=None):
    spacing = ''
    if space_before or space_after:
        sp = '<w:spacing'
        if space_before: sp += f' w:before="{space_before}"'
        if space_after: sp += f' w:after="{space_after}"'
        spacing = sp + '/>'
    rPr = '<w:rFonts w:hint="cs"/>'
    if bold: rPr += '<w:b/><w:bCs/>'
    if size: rPr += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    if color: rPr += f'<w:color w:val="{color}"/>'
    rPr += '<w:rtl/>'
    t_attr = ' xml:space="preserve"' if ' ' in text else ''
    return (f'<w:p><w:pPr><w:bidi/>{spacing}<w:rPr><w:rtl/></w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rPr}</w:rPr><w:t{t_attr}>{text}</w:t></w:r></w:p>')

def heb_bullet(text, size=24, space_after=80):
    rPr = f'<w:rFonts w:hint="cs"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/><w:rtl/>'
    t_attr = ' xml:space="preserve"' if ' ' in text else ''
    return (f'<w:p><w:pPr>'
            f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
            f'<w:bidi/><w:spacing w:after="{space_after}"/>'
            f'<w:rPr><w:rtl/></w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rPr}</w:rPr><w:t{t_attr}>{text}</w:t></w:r></w:p>')

def heb_empty():
    return '<w:p><w:pPr><w:bidi/><w:rPr><w:rtl/></w:rPr></w:pPr></w:p>'
```

### RTL Bullets — numbering.xml:
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x2022;"/>
      <w:lvlJc w:val="right"/>
      <w:pPr>
        <w:bidi/>
        <w:ind w:right="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:hint="cs"/>
        <w:rtl/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>
```

Add to `word/_rels/document.xml.rels`:
```xml
<Relationship Id="rId99" Type=".../numbering" Target="numbering.xml"/>
```
Add to `[Content_Types].xml`:
```xml
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
```

### Tables in Hebrew docx:
- Set `bidiVisual` on the table: `<w:tblPr><w:bidiVisual/></w:tblPr>`
- First column in the XML = rightmost column visually
- Header row cells use the same `<w:bidi/>` + `<w:rtl/>` pattern

---

## PowerPoint (.pptx)

### Text boxes:
Always set both `align: "right"` AND `rtlMode: true` on every Hebrew text box:
```javascript
slide.addText("הטקסט כאן", {
  fontFace: "Arial", fontSize: 24,
  align: "right", rtlMode: true
});
```

### Bullets, tables, timelines, grids — REQUIRED POST-PROCESSING

PptxGenJS does not correctly handle RTL bullets or RTL column order. After writing the pptx, run this Python fix:

```python
import zipfile, os, re

def fix_rtl_pptx(src, dst):
    os.makedirs('/tmp/pptx_fix', exist_ok=True)
    with zipfile.ZipFile(src, 'r') as z:
        z.extractall('/tmp/pptx_fix')

    slides_dir = '/tmp/pptx_fix/ppt/slides'
    for fname in os.listdir(slides_dir):
        if fname.endswith('.xml') and fname.startswith('slide'):
            fpath = os.path.join(slides_dir, fname)
            content = open(fpath, encoding='utf-8').read()
            content = fix_slide(content)
            open(fpath, 'w', encoding='utf-8').write(content)

    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for root, dirs, files in os.walk('/tmp/pptx_fix'):
            for file in files:
                fp = os.path.join(root, file)
                zout.write(fp, os.path.relpath(fp, '/tmp/pptx_fix'))

def fix_slide(content):
    # 1. Set rtlCol="1" on all text body containers
    content = re.sub(r'rtlCol="0"', 'rtlCol="1"', content)

    # 2. Fix bullet paragraphs: add rtl="1", swap marL->marR, fix indent direction
    def fix_para_with_bullet(m):
        para = m.group(0)
        if 'buChar' not in para and 'buAutoNum' not in para:
            return para
        def fix_pPr(pm):
            ppr = pm.group(0)
            if 'rtl="1"' not in ppr:
                ppr = ppr.replace('<a:pPr ', '<a:pPr rtl="1" ')
            ppr = re.sub(r'marL="(\d+)"', lambda x: f'marR="{x.group(1)}"', ppr)
            ppr = re.sub(r' indent="-(\d+)"', lambda x: f' indent="{x.group(1)}"', ppr)
            return ppr
        para = re.sub(r'<a:pPr[^/]*/>', fix_pPr, para)
        return para
    content = re.sub(r'<a:p>.*?</a:p>', fix_para_with_bullet, content, flags=re.DOTALL)

    # 3. Set Hebrew language on runs with Hebrew text
    def fix_run_lang(m):
        run = m.group(0)
        if re.search(r'[\u0590-\u05FF]', run):
            run = re.sub(r'lang="en-US"', 'lang="he-IL"', run)
        return run
    content = re.sub(r'<a:r>.*?</a:r>', fix_run_lang, content, flags=re.DOTALL)

    return content
```

### Tables in Hebrew pptx:
- Build columns in **reverse order** in the code (last column first in JS), OR
- After generation, reverse the column order in XML via post-processing
- Header row: rightmost cell = first/primary column

### Timelines and grids:
- First item (earliest / most important) goes on the RIGHT
- Build items in the array right-to-left: `[item_last, ..., item_first]` with x-positions reversed
- Example for a 3-step timeline:
```javascript
const items = ['שלב א', 'שלב ב', 'שלב ג']; // right to left
const startX = 11.0; // start from right
items.forEach((item, i) => {
  slide.addText(item, {
    x: startX - (i * 3.5), y: 3.0, w: 3.0, h: 1.0,
    align: "right", rtlMode: true, fontFace: "Arial"
  });
});
```

---

## HTML / PDF

```html
<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
  <meta charset="UTF-8">
  <link href="https://fonts.googleapis.com/css2?family=Heebo:wght@400;700&display=swap" rel="stylesheet">
  <style>
    body { font-family: 'Heebo', Arial, sans-serif; direction: rtl; text-align: right; }
    ul, ol { padding-right: 1.5em; padding-left: 0; }
    table { direction: rtl; }
    th, td { text-align: right; }
  </style>
</head>
```

Tables in HTML: `direction: rtl` on the `<table>` element automatically puts the first column on the right.

For PDF: `from weasyprint import HTML; HTML(string=html).write_pdf("out.pdf")`

---

## Checklist

**docx:**
- [ ] Using real Hebrew Word template as base
- [ ] `<w:bidi/>` (no attributes) on every paragraph
- [ ] `<w:rFonts w:hint="cs"/>` + `<w:rtl/>` on every run
- [ ] Tables: `<w:bidiVisual/>` in tblPr
- [ ] Bullets: numbering.xml with `<w:lvlJc w:val="right"/>` and `<w:bidi/>`

**pptx:**
- [ ] `align: "right"` + `rtlMode: true` on all text boxes
- [ ] Run `fix_rtl_pptx()` post-processing on every pptx
- [ ] Tables/grids/timelines: first item on the RIGHT

**All formats:**
- [ ] No em-dash `—` — only regular hyphen `-`
- [ ] Font is Hebrew-compatible (Arial for docx/pptx, Heebo for HTML)
