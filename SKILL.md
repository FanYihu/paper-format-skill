---
name: paper-format
description: >
  Use this skill whenever a user wants to analyze a template and apply its formatting
  rules to their own .docx paper. Triggers include: "按这个PDF格式排版我的论文",
  "提取论文格式规则", "把我的论文改成这个格式", "分析模板PDF的排版", "套用论文格式",
  "apply formatting from PDF", "match this paper's format", "参考这篇论文的格式",
  "仿照这个排版", "按格式要求修改我的论文". Also triggers when a user uploads a
  template (PDF, docx, or text requirements) alongside their own .docx and asks to
  reformat. Always use this skill — do NOT handle multi-step format workflows freehand.
  Covers: heading hierarchy & fonts, page margins, line/paragraph spacing,
  header/footer/page numbers, references style, figure/table captions.
---

# Paper Format Skill

Three-phase workflow:
**Phase 1 Analyze template → Phase 2 Apply + deep-clear overrides → Phase 3 Verify**

Always show the user the extracted rules table and get confirmation before Phase 2.
Always run Phase 3 verification after Phase 2 and present the report to the user.

---

## Phase 1 — Analyze template (PDF / docx / text requirements)

Detect the template type first, then use the matching extraction path.

### 1A — Template is a PDF (already-formatted paper sample)

```bash
pip install pdfplumber --break-system-packages -q
pdfinfo "$TEMPLATE_PDF"
pdftotext -layout -f 1 -l 4 "$TEMPLATE_PDF" /tmp/sample.txt && cat /tmp/sample.txt
```

If text is empty or garbled (scanned PDF), rasterize for visual inspection:
```bash
pdftoppm -jpeg -r 150 -f 1 -l 2 "$TEMPLATE_PDF" /tmp/tpl_page
ls /tmp/tpl_page*.jpg   # then view each image
```

Run deep metric extraction:
```python
import pdfplumber, json
from collections import Counter

results = {}
with pdfplumber.open("$TEMPLATE_PDF") as pdf:
    p0 = pdf.pages[0]
    w, h = p0.width, p0.height
    results["page"] = {
        "width_mm":  round(w * 0.352778, 1),
        "height_mm": round(h * 0.352778, 1),
    }
    words = p0.extract_words()
    if words:
        xs  = [wd["x0"] for wd in words]
        x1s = [wd["x1"] for wd in words]
        ys  = [wd["top"] for wd in words]
        results["margins_mm"] = {
            "left":   round(min(xs)       * 0.352778, 1),
            "right":  round((w-max(x1s))  * 0.352778, 1),
            "top":    round(min(ys)        * 0.352778, 1),
            "bottom": round((h-max([wd["bottom"] for wd in words])) * 0.352778, 1),
        }
    all_chars = []
    for pg in pdf.pages[:4]:
        all_chars.extend(pg.chars)
    font_counter = Counter()
    for ch in all_chars:
        key = (ch.get("fontname","?"), round(ch.get("size", 0), 1))
        font_counter[key] += 1
    results["fonts_ranked"] = [
        {"name": k[0], "size_pt": k[1], "char_count": v}
        for k, v in font_counter.most_common(20)
    ]
    tops = sorted(set(round(ch["top"], 0) for ch in all_chars))
    gaps = [round(tops[i+1]-tops[i], 1) for i in range(len(tops)-1)
            if 0 < tops[i+1]-tops[i] < 40]
    results["line_gap_pt_freq"] = Counter(gaps).most_common(8)
    hdr = [ch.get("text","") for ch in p0.chars if ch["top"] < h * 0.08]
    ftr = [ch.get("text","") for ch in p0.chars if ch["bottom"] > h * 0.92]
    results["header_text_sample"] = "".join(hdr)[:120]
    results["footer_text_sample"] = "".join(ftr)[:120]
    caps = []
    for pg in pdf.pages[:6]:
        for line in (pg.extract_text() or "").splitlines():
            l = line.strip()
            if l and any(l.startswith(p) for p in ["图","Fig","Table","表"]):
                caps.append(l[:80])
    results["caption_samples"] = caps[:6]
    refs, in_refs = [], False
    for pg in pdf.pages[-3:]:
        for line in (pg.extract_text() or "").splitlines():
            l = line.strip()
            if any(kw in l for kw in ["参考文献","References","REFERENCES"]):
                in_refs = True
            if in_refs and l:
                refs.append(l[:80])
    results["reference_samples"] = refs[:8]
print(json.dumps(results, indent=2, ensure_ascii=False))
```

**Decision rules for synthesis:**
- Body font = most frequent (fontname, size) pair, excluding header/footer zone chars
- Heading fonts = pairs significantly larger than body, grouped by size descending
- Line spacing multiple ≈ most common line_gap_pt / body_font_size_pt
- Reference style: `[1]` pattern → GB/T 7714 or IEEE; `Author (Year)` → APA

### 1B — Template is a .docx (another formatted paper)

Read styles directly — far more accurate than PDF inference:
```python
from docx import Document
from docx.shared import Pt, Mm

doc = Document("$TEMPLATE_DOCX")

results = {}

# Page geometry
s = doc.sections[0]
results["page"] = {
    "width_mm":  round(s.page_width.mm,  1),
    "height_mm": round(s.page_height.mm, 1),
}
results["margins_mm"] = {
    "top":    round(s.top_margin.mm,    1),
    "bottom": round(s.bottom_margin.mm, 1),
    "left":   round(s.left_margin.mm,   1),
    "right":  round(s.right_margin.mm,  1),
}

# Read key styles
def read_style(doc, name):
    try:
        st = doc.styles[name]
        f  = st.font
        pf = st.paragraph_format
        return {
            "font_name":    f.name,
            "font_size_pt": f.size.pt if f.size else None,
            "bold":         f.bold,
            "alignment":    str(pf.alignment),
            "line_spacing": str(pf.line_spacing),
            "space_before_pt": pf.space_before.pt if pf.space_before else 0,
            "space_after_pt":  pf.space_after.pt  if pf.space_after  else 0,
        }
    except Exception as e:
        return {"error": str(e)}

for sname in ["Normal", "Body Text", "Heading 1", "Heading 2", "Heading 3",
              "参考文献条目", "章标题"]:
    results[sname] = read_style(doc, sname)

# Header/footer text
hdr_texts = [p.text for p in doc.sections[0].header.paragraphs if p.text.strip()]
ftr_texts = [p.text for p in doc.sections[0].footer.paragraphs if p.text.strip()]
results["header_texts"] = hdr_texts
results["footer_texts"]  = ftr_texts

import json
print(json.dumps(results, indent=2, ensure_ascii=False))
```

### 1C — Template is text (format requirements document / spec)

When the user pastes or uploads a plain-text format specification, use Claude to
parse it into format_rules.json directly. Ask Claude (inner call or direct reasoning)
to extract every measurable rule and produce the canonical JSON below.
Prompt template:
```
以下是论文格式要求文字说明，请将其解析为结构化JSON，字段名与结构严格按照 format_rules.json
模板。无法确定的字段设为 null，不要猜测。

[用户的格式要求文字]
```

---

### Step 1C-end / 1A-end / 1B-end: Produce format_rules.json

Fill the canonical rule object from whichever extraction path was used:

```json
{
  "page": { "width_mm": 210, "height_mm": 297 },
  "margins_mm": { "top": 25, "bottom": 25, "left": 30, "right": 25 },
  "body_text": {
    "font_name_latin": "Times New Roman",
    "font_name_eastasia": "宋体",
    "font_size_pt": 10.5,
    "line_spacing": "1.5x",
    "paragraph_spacing_before_pt": 0,
    "paragraph_spacing_after_pt": 0,
    "alignment": "justify",
    "first_line_indent_chars": 2
  },
  "headings": {
    "h1": { "font_name_latin": "Times New Roman", "font_name_eastasia": "黑体",
            "font_size_pt": 14, "bold": true, "alignment": "left",
            "space_before_pt": 24, "space_after_pt": 12 },
    "h2": { "font_name_latin": "Times New Roman", "font_name_eastasia": "黑体",
            "font_size_pt": 11, "bold": true, "alignment": "left",
            "space_before_pt": 18, "space_after_pt": 6 },
    "h3": { "font_name_latin": "Times New Roman", "font_name_eastasia": "黑体",
            "font_size_pt": 11, "bold": true, "alignment": "left",
            "space_before_pt": 12, "space_after_pt": 6 }
  },
  "header_footer": {
    "header_text": "",
    "header_font_size_pt": 9,
    "header_alignment": "center",
    "footer_page_number": true,
    "footer_page_number_position": "center",
    "footer_font_size_pt": 9,
    "first_page_different": false
  },
  "references": {
    "style": "GB/T 7714",
    "font_size_pt": 9,
    "line_spacing": "1.0x",
    "hanging_indent_mm": 7,
    "space_after_pt": 3
  },
  "figures_tables": {
    "figure_caption_position": "below",
    "table_caption_position": "above",
    "caption_font_size_pt": 9,
    "caption_bold": false,
    "caption_alignment": "center",
    "figure_prefix": "图",
    "table_prefix": "表"
  },
  "_protected_styles": ["HTML Preformatted", "Source Code"]
}
```

Save to `/tmp/format_rules.json`.

**Show the user a confirmation table and wait for approval before Phase 2.**

---

## Phase 2 — Apply + Deep-clear overrides

Only proceed after user confirms the rules table.

### Step 2.1: Unpack

```bash
pip install python-docx --break-system-packages -q
python /mnt/skills/public/docx/scripts/office/unpack.py "$USER_DOCX" /tmp/unpacked/
```

### Step 2.2: Deep-clear run-level and paragraph-level direct overrides

**This is the most critical step — without it, style changes have no effect.**

Word's formatting cascade (lowest to highest priority):
  `style definition` → `paragraph direct pPr` → `run direct rPr`

Overrides at the paragraph and run level silently win over everything in the style
sheet. They must be stripped before (or alongside) style edits.

```python
from lxml import etree
import copy, json

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

rules = json.load(open("/tmp/format_rules.json"))
protected_styles = set(rules.get("_protected_styles",
                                  ["HTML Preformatted", "Source Code"]))

# Tags to strip from run-level rPr (font, size, bold, italic, color overrides)
RUN_CLEAR_TAGS = {
    f"{{{W}}}rFonts", f"{{{W}}}sz", f"{{{W}}}szCs",
    f"{{{W}}}b",      f"{{{W}}}bCs",
    f"{{{W}}}i",      f"{{{W}}}iCs",
    f"{{{W}}}color",  f"{{{W}}}kern",
}

# Tags to strip from paragraph-level pPr (spacing, indent direct overrides)
PARA_CLEAR_TAGS = {
    f"{{{W}}}spacing",
    f"{{{W}}}ind",
    f"{{{W}}}jc",
}

tree = etree.parse("/tmp/unpacked/word/document.xml")
root = tree.getroot()

cleared_runs  = 0
cleared_paras = 0
protected     = 0

for p_el in root.iter(f"{{{W}}}p"):
    # Determine paragraph style
    pPr = p_el.find(f"{{{W}}}pPr")
    style_name = ""
    if pPr is not None:
        pStyle = pPr.find(f"{{{W}}}pStyle")
        if pStyle is not None:
            style_name = pStyle.get(f"{{{W}}}val", "")

    # Check if this paragraph type is protected (code blocks etc.)
    is_protected = any(ps.lower().replace(" ","") in style_name.lower().replace(" ","")
                       for ps in protected_styles)
    if style_name in {"HTMLPreformatted", "SourceCode", "HTML Preformatted", "Source Code"}:
        is_protected = True

    if is_protected:
        protected += 1
        continue

    # ── Clear paragraph-level direct overrides ─────────────────────────
    if pPr is not None:
        for tag in list(PARA_CLEAR_TAGS):
            for el in pPr.findall(tag):
                pPr.remove(el)
                cleared_paras += 1

    # ── Clear run-level direct overrides ───────────────────────────────
    for r_el in p_el.findall(f"{{{W}}}r"):
        rPr = r_el.find(f"{{{W}}}rPr")
        if rPr is None:
            continue
        for tag in list(RUN_CLEAR_TAGS):
            for el in rPr.findall(tag):
                rPr.remove(el)
                cleared_runs += 1
        # Remove empty rPr
        if len(rPr) == 0:
            r_el.remove(rPr)

tree.write("/tmp/unpacked/word/document.xml",
           xml_declaration=True, encoding="utf-8", standalone=True)
print(f"✅ Cleared {cleared_runs} run overrides, {cleared_paras} para overrides, "
      f"skipped {protected} protected paragraphs")
```

### Step 2.3: Edit styles.xml — update style definitions

Edit `/tmp/unpacked/word/styles.xml` to update Normal, Body Text, Heading 1/2/3,
章标题, 参考文献条目 style definitions. For each style, set:
- `<w:rFonts>` with correct ascii/hAnsi/eastAsia values
- `<w:sz>` and `<w:szCs>` (half-points: 10pt → val="20")
- `<w:spacing>` with before/after/line/lineRule
- `<w:jc>` for alignment
- `<w:ind>` for first-line indent where needed

Reference values (half-points for sz, twentieths-of-a-point for spacing):
```
10pt  → sz val="20"     |  1.5x line → line="360" lineRule="auto"
10.5pt → sz val="21"    |  1.0x line → line="240" lineRule="auto"
11pt  → sz val="22"     |  single    → lineRule="exact" line="240"
14pt  → sz val="28"     |  24pt space → spacing before="480"
9pt   → sz val="18"     |  12pt space → spacing before="240"
```

### Step 2.4: Add/update header and footer

Create `/tmp/unpacked/word/header1.xml` and `/tmp/unpacked/word/footer1.xml`,
add relationships to `_rels/document.xml.rels`, update content types, and
add `<w:headerReference>` / `<w:footerReference>` to the `<w:sectPr>`.

Footer page number uses a `<w:fldChar>` / `<w:instrText xml:space="preserve"> PAGE </w:instrText>` field run.

### Step 2.5: Pack

```bash
python /mnt/skills/public/docx/scripts/office/pack.py \
    /tmp/unpacked/ /tmp/output_formatted.docx \
    --original "$USER_DOCX"
```

---

## Phase 3 — Verify (mandatory, always run after Phase 2)

### Step 3.1: Machine verification — read back and compare

```python
from docx import Document
from docx.shared import Pt, Mm
import json

rules  = json.load(open("/tmp/format_rules.json"))
doc    = Document("/tmp/output_formatted.docx")

PASS = "✅"; FAIL = "❌"; WARN = "⚠️"
report = []

def check(label, got, expected, tol=0.5):
    """Compare numeric values within tolerance."""
    try:
        ok = abs(float(got) - float(expected)) <= tol
        status = PASS if ok else FAIL
    except Exception:
        ok = str(got) == str(expected)
        status = PASS if ok else FAIL
    report.append({"item": label, "expected": expected, "actual": got, "status": status})
    return ok

# ── Page & margins ──────────────────────────────────────────────────
s = doc.sections[0]
check("纸张宽度(mm)",  round(s.page_width.mm,  1), rules["page"]["width_mm"])
check("纸张高度(mm)",  round(s.page_height.mm, 1), rules["page"]["height_mm"])
m = rules["margins_mm"]
check("上边距(mm)",  round(s.top_margin.mm,    1), m["top"])
check("下边距(mm)",  round(s.bottom_margin.mm, 1), m["bottom"])
check("左边距(mm)",  round(s.left_margin.mm,   1), m["left"])
check("右边距(mm)",  round(s.right_margin.mm,  1), m["right"])

# ── Normal / Body Text style ────────────────────────────────────────
body = rules["body_text"]
for sname in ["Normal", "Body Text"]:
    try:
        st = doc.styles[sname]
        check(f"{sname} 字号(pt)", st.font.size.pt if st.font.size else "None",
              body["font_size_pt"])
        ls = st.paragraph_format.line_spacing
        expected_ls = float(body["line_spacing"].replace("x","")) if "x" in body["line_spacing"] else None
        if expected_ls and ls:
            check(f"{sname} 行距倍数", round(float(ls), 2), expected_ls)
        sb = st.paragraph_format.space_before
        check(f"{sname} 段前(pt)", round(sb.pt, 1) if sb else 0,
              body.get("paragraph_spacing_before_pt", 0))
        sa = st.paragraph_format.space_after
        check(f"{sname} 段后(pt)", round(sa.pt, 1) if sa else 0,
              body.get("paragraph_spacing_after_pt", 0))
    except Exception as e:
        report.append({"item": f"{sname} 读取失败", "expected": "-", "actual": str(e), "status": WARN})

# ── Heading styles ───────────────────────────────────────────────────
H_MAP = {"h1":"Heading 1","h2":"Heading 2","h3":"Heading 3"}
for level, sname in H_MAP.items():
    if level not in rules.get("headings", {}):
        continue
    h = rules["headings"][level]
    try:
        st = doc.styles[sname]
        check(f"{sname} 字号(pt)",
              st.font.size.pt if st.font.size else "None", h["font_size_pt"])
    except Exception as e:
        report.append({"item": f"{sname} 读取失败", "expected": "-",
                       "actual": str(e), "status": WARN})

# ── Header & footer ──────────────────────────────────────────────────
hf = rules.get("header_footer", {})
hdr_texts = [p.text.strip() for p in doc.sections[0].header.paragraphs if p.text.strip()]
ftr_paras  = [p for p in doc.sections[0].footer.paragraphs]
has_header = len(hdr_texts) > 0
has_footer = len(ftr_paras) > 0
report.append({"item": "页眉存在", "expected": bool(hf.get("header_text","")) or True,
               "actual": has_header, "status": PASS if has_header else FAIL})
report.append({"item": "页脚/页码存在", "expected": hf.get("footer_page_number", True),
               "actual": has_footer, "status": PASS if has_footer else FAIL})

# ── Run-level override check ─────────────────────────────────────────
# Count remaining run-level font/size overrides in body paragraphs
from lxml import etree
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PROTECTED = set(rules.get("_protected_styles", ["HTML Preformatted", "Source Code"]))
remaining_run_overrides = 0
remaining_para_overrides = 0

doc2 = Document("/tmp/output_formatted.docx")
for para in doc2.paragraphs:
    if para.style.name in PROTECTED:
        continue
    pPr = para._p.find(f"{{{W}}}pPr")
    if pPr is not None:
        for tag in [f"{{{W}}}spacing", f"{{{W}}}ind"]:
            if pPr.find(tag) is not None:
                remaining_para_overrides += 1
    for r in para._p.findall(f"{{{W}}}r"):
        rPr = r.find(f"{{{W}}}rPr")
        if rPr is not None:
            for tag in [f"{{{W}}}sz", f"{{{W}}}rFonts"]:
                if rPr.find(tag) is not None:
                    remaining_run_overrides += 1
                    break

run_status  = PASS if remaining_run_overrides  == 0 else FAIL
para_status = PASS if remaining_para_overrides == 0 else FAIL
report.append({"item": "残留run字体/字号覆盖段落数",
               "expected": 0, "actual": remaining_run_overrides,  "status": run_status})
report.append({"item": "残留段落直接spacing/ind覆盖数",
               "expected": 0, "actual": remaining_para_overrides, "status": para_status})

# ── Print report table ───────────────────────────────────────────────
passed = sum(1 for r in report if r["status"] == PASS)
failed = sum(1 for r in report if r["status"] == FAIL)
warned = sum(1 for r in report if r["status"] == WARN)

print(f"\n{'='*60}")
print(f"格式核查报告   ✅{passed} 通过  ❌{failed} 失败  ⚠️{warned} 警告")
print(f"{'='*60}")
print(f"{'核查项':<28} {'期望值':<12} {'实际值':<12} 状态")
print(f"{'-'*60}")
for r in report:
    print(f"{r['item']:<28} {str(r['expected']):<12} {str(r['actual']):<12} {r['status']}")
print(f"{'='*60}")
```

### Step 3.2: Manual confirmation checklist

Always present this checklist to the user alongside the auto report:

```
📋 建议在 Word 中手动确认的项目：

□ 封面/首页格式是否有特殊要求（独立样式或不同页边距）
□ 摘要、关键词段落字体是否符合要求
□ 图表编号是否连贯（按 Ctrl+A → F9 刷新域）
□ 参考文献每条悬挂缩进是否一致
□ 代码块字体（Menlo/Courier）是否保持不变（应未修改）
□ 表格内文字字体/字号是否需要单独调整
□ 脚注/尾注格式（如有）
□ 页眉在每页是否正确显示
```

---

## Step 4: Deliver

```bash
python /mnt/skills/public/docx/scripts/office/validate.py /tmp/output_formatted.docx
cp /tmp/output_formatted.docx /mnt/user-data/outputs/formatted_paper.docx
```

Call `present_files` with the output path.

---

## Edge Cases

| Situation | Handling |
|-----------|----------|
| Scanned PDF (no text layer) | Rasterize with pdftoppm → view images → ask user to confirm measurements |
| Embedded Chinese fonts | Map: 宋体→SimSun, 黑体→SimHei, 楷体→KaiTi, 仿宋→FangSong |
| Complex tables in .docx | Apply page/margin/body/header; skip table cell font unless asked |
| Header/footer contains logo | Extract image, re-insert as InlineImage in python-docx header |
| .docx has no Heading styles | Warn user: Heading 1/2/3 styles required; offer font-size heuristic fallback |
| Cross-reference fields | Inform user: Ctrl+A → F9 in Word to refresh all field numbers |
| Verification shows residual overrides | Re-run Step 2.2 with wider PARA_CLEAR_TAGS; check for inline styles in tables |
| Text format requirements conflict | Flag conflicts to user and ask which takes precedence |

---

## Final Output to User (always all four)

1. **Rules confirmation table** — Phase 1 output, wait for user approval
2. **Formatted .docx** — via `present_files`
3. **Auto verification report** — ✅/❌ table from Phase 3.1
4. **Manual confirmation checklist** — Phase 3.2 checklist for Word
