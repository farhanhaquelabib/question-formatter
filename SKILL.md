---
name: phoenix-image-to-question-converter
description: >-
  MANDATORY AUTOTRIGGER: Use this skill whenever images, photos, camera captures, screenshots, scans, or PDFs containing questions, test papers, or math problem sets are uploaded or provided to create, transcribe, solve, rephrase, or format questions into exam papers. MUST activate automatically whenever images/PDFs of questions are provided, REGARDLESS of whether the user explicitly mentions this skill name or not. Enforces strict typography (Cambria Math 11pt bold for English questions, no italics anywhere), plain text options with ruler tab stops (NEVER use tables for options), exact 0.75" margins all around (Top: 0.75", Bottom: 0.75", Left: 0.75", Right: 0.75"), proper OMML math equations (fractions, vinculum square roots, auto-scaling brackets), and 10 questions per page budgeting. English questions only by default (NO Bangla translation unless explicitly requested).
---

# Phoenix Image to Question Converter

This skill establishes the standardized workflow, typography, layout, OMML math equation formatting (fractions, auto-scaling bracket delimiters, and radicals), zero-italic enforcement, margin specifications, and file export procedures for converting images of questions into professionally formatted exam papers and solution sheets in DOCX and PDF formats.

---

## 1. Triggering & Activation Rules (MANDATORY)

1. **Automatic Activation on Images/PDFs:**
   - **Always trigger automatically** whenever the user uploads, attaches, or provides images (photos, screenshots, camera captures, scans) or PDFs containing exam questions, test papers, or math problem sets with requests like:
     - *"format these questions"*
     - *"create questions from these pictures"*
     - *"convert questions to docx/pdf"*
     - *"transcribe the questions"*
     - *"solve these questions"*
     - *"rephrase/re-digit these questions"*
   - **No Explicit Mention Required:** Activate this skill **regardless of whether the user explicitly mentions "Phoenix Image to Question Converter", "the skill", or any skill name**. The presence of images of questions with a formatting/creation prompt is the sole trigger required.

2. **Batch Uploads & Waiting Condition:**
   - If the user indicates they are uploading pictures in multiple attempts, acknowledge each batch, list the detected pages, and **wait** until the user explicitly says `"Done, you may start"` (or equivalent start prompt).
3. **Duplicate Detection:**
   - Always track page numbers and question ranges from every uploaded image.
   - If a duplicate image/page is uploaded, **immediately alert the user** specifying which page was duplicated and which images contained it.
4. **Question Filtering:**
   - Strictly respect user filtering criteria (e.g., if the user says *"Only count the math questions"*, exclude all non-math questions such as general knowledge or English questions).

---

## 2. Content Preservation & Language Rules

1. **Preserve Original Wording & Digits by Default (CRITICAL):**
   - **Do NOT** change the question wording or digits unless explicitly instructed by the user.
   - If the user asks to format, create, or transcribe questions without mentioning digit changes, **keep the question wording and numerical values exactly as they are** in the source material.
   - Only alter digits, parameters, or rephrase if the user explicitly requests: *"change the wording and digits"* or *"make fresh questions"*.
2. **When Rephrasing/Re-digiting is Explicitly Requested:**
   - Ensure new numbers remain mathematically consistent, physically plausible, and yield clean, neat answers (integers or simple terminating decimals).
3. **Language Policy — English Only by Default (CRITICAL):**
   - **Default:** Generate questions in **English only**.
     - Structure: `[Question Number]. [English Question Text]`
     - **Do NOT** include Bangla translations unless the user explicitly requests them.
   - **If (and only if) Prompted for Bangla:**
     - Include the Bangla translation inside parentheses immediately following the English question:
       `[Question Number]. [English Question Text] ([Bangla Translation])`
     - Typography: **Nirmala UI, 9 pt, Regular (Unbolded)**.
     - **English Digits Only:** In the Bangla text, all numbers must use standard English digits (`0–9`). **No Bengali digits (`০–৯`) are allowed.**
     - **Equation Edits in Bangla Text:** Any fractions (e.g. 1/10), roots (e.g. √3), subscripts, or superscripts (e.g. x²) involving English digits or variables inside the Bangla translation must use proper native OMML equation formatting.
4. **Four Options (MCQs):**
   - Provide 4 distinct options (`A.`, `B.`, `C.`, `D.`) per question.
   - Ensure only one option is unambiguously correct, with realistic distractors for the remaining three.

---

## 3. Strict Typography & Zero-Italics Policy

- **Question Numbering:** Consecutive two digits with leading zeros (`01.`, `02.`, ..., `40.`) in **bold Cambria Math, 11 pt**.
- **No Italic Words in the Final File (CRITICAL):**
  - **Zero Italics Policy:** Absolutely no italic words, letters, or math variables anywhere in the document.
  - In normal text runs: `run.italic = False`, `<w:i w:val="0"/>`, `<w:iCs w:val="0"/>`.
  - In Office Math (OMML): Word's math engine defaults letters ($x, y, a, b, P, r$, etc.) to italics. **You MUST add `<m:nor/>` (math normal style) inside every `<m:rPr>`** so that math variables remain completely upright and non-italic.
- **English Question Typography (BOLD):**
  - Font: **Cambria Math**
  - Size: **11 pt**
  - Style: **Bold (`bold=True`)**, strictly un-italicized (`run.italic=False`)
  - Color: Black (`#000000`)
- **Options Typography (REGULAR):**
  - Font: **Cambria Math**
  - Size: **11 pt**
  - Style: **Regular / Unbolded (`bold=False`)**, strictly un-italicized (`run.italic=False`)
  - Prefix: `A. `, `B. `, `C. `, `D. `

---

## 4. Fraction, Radical & Bracket Equation Formatting (OMML)

1. **No Spaces Around Slashes:**
   - For all inline ratios or units (e.g., `km/h`, `m/min`, `x:y`), **NEVER put spaces before or after `/`**.
2. **Proper Vertical Equation Fractions:**
   - Whenever fractions appear as mathematical expressions (e.g., 1/10, 3/4, 5/7, 7/8, 8/11) or **algebraic divisions** (e.g., `6000/300`, `25/1.25`, `360/0.90`, `4250/0.85`, `580/29`, `750/15`, `W/1.25`), convert them into native **Office Math (OMML)** equations with a stacked vertical bar:
   ```python
   def add_omml_fraction(paragraph, num_content, den_content, bold=False):
       num_xml = _build_omml_element(num_content, bold=bold)
       den_xml = _build_omml_element(den_content, bold=bold)
       omml_str = (
           f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<m:f>'
           f'<m:fPr><m:type m:val="bar"/></m:fPr>'
           f'<m:num>{num_xml}</m:num>'
           f'<m:den>{den_xml}</m:den>'
           f'</m:f>'
           f'</m:oMath>'
       )
       paragraph._p.append(parse_xml(omml_str))
   ```
3. **Proper Equation Radicals (Square Roots):**
   - All square roots in math must be converted into native OMML `<m:rad>` format so that the radical bar (vinculum) spans across the entire number or expression (e.g., √3, √400):
   ```python
   def add_omml_sqrt(paragraph, expr_text, bold=False):
       bold_tag = "<w:b/>" if bold else ""
       omml_str = (
           f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<m:rad>'
           f'<m:radPr><m:degHide m:val="1"/></m:radPr>'
           f'<m:deg/>'
           f'<m:e><m:r><m:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><m:nor/>{bold_tag}</m:rPr><m:t>{expr_text}</m:t></m:r></m:e>'
           f'</m:rad>'
           f'</m:oMath>'
       )
       paragraph._p.append(parse_xml(omml_str))
   ```
4. **Auto-Scaling Bracket Delimiters (`<m:d>`):**
   - Whenever an equation part, fraction, or algebraic expression is inside parentheses `( )` (e.g. (√3/4), (22/7), (3/5), (6/π)², (x + 1/x)²), **include the brackets inside the equation using the OMML `<m:d>` element**:
   ```python
   def add_omml_paren_frac(paragraph, num_content, den_content, bold=False, sup=None):
       num_xml = _build_omml_element(num_content, bold=bold)
       den_xml = _build_omml_element(den_content, bold=bold)
       frac_xml = f'<m:f><m:fPr><m:type m:val="bar"/></m:fPr><m:num>{num_xml}</m:num><m:den>{den_xml}</m:den></m:f>'
       delim_xml = f'<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>{frac_xml}</m:e></m:d>'
       if sup:
           sup_xml = _build_omml_element(sup, bold=bold)
           omml_str = f'<m:oMath ...><m:sSup><m:e>{delim_xml}</m:e><m:sup>{sup_xml}</m:sup></m:sSup></m:oMath>'
       else:
           omml_str = f'<m:oMath ...>{delim_xml}</m:oMath>'
       paragraph._p.append(parse_xml(omml_str))
   ```
   - This ensures the bracket height automatically stretches and scales to match the exact vertical height of the stacked fraction or expression.

---

## 5. Document Margins & Page Layout Specifications

1. **Document Margins (CRITICAL — STRICT 0.75" ALL AROUND):**
   - **Top Margin:** `0.75 in` (1080 dxa)
   - **Bottom Margin:** `0.75 in` (1080 dxa)
   - **Left Margin:** `0.75 in` (1080 dxa)
   - **Right Margin:** `0.75 in` (1080 dxa)
   - **Page Size:** Standard Letter (`8.5 in` × `11.0 in`)
   - **Printable Width:** `7.0 in` ($8.5 - 0.75 - 0.75 = 7.0$ inches / 10,080 dxa).

2. **Section Header:**
   - 1-row, 2-column borderless table across full printable width (`7.0 in`):
     - Left cell (`4.5 in`): `Section 01: [Subject Name]` (Bold, 12 pt Cambria Math)
     - Right cell (`2.5 in`): `Total Questions: [Count]` (Bold, 12 pt Cambria Math, right-aligned)

3. **MCQ Options Layout — NO TABLES (STRICT RULE):**
   - **Never use tables for options:** Whatever you do, do NOT use tables for MCQ options.
   - **Plain Text with Ruler Tab Stops:** Write the 4 options (`A.`, `B.`, `C.`, `D.`) as a single paragraph using OpenXML ruler tab stops (`<w:tabs>`) and `<w:tab/>` spacing.
   - **Tab Stop Positions for 0.75" Margins (7.0" Printable Width):**
     - **Option A:** Margin start (`0 in` / `0 dxa`)
     - **Option B:** `1.750 in` (`2520 dxa`)
     - **Option C:** `3.500 in` (`5040 dxa`)
     - **Option D:** `5.250 in` (`7560 dxa`)
   - **XML Implementation:**
     ```python
     TAB_STOPS = [2520, 5040, 7560]

     def set_paragraph_tab_stops(p, positions_dxa=TAB_STOPS):
         pPr = p._p.get_or_add_pPr()
         tabs = parse_xml(f'<w:tabs {nsdecls("w")}/>')
         for pos in positions_dxa:
             tab_elem = parse_xml(f'<w:tab {nsdecls("w")} w:val="left" w:pos="{pos}"/>')
             tabs.append(tab_elem)
         pPr.append(tabs)

     def add_tab_run(p):
         run = p.add_run()
         run.element.append(parse_xml(f'<w:tab {nsdecls("w")}/>'))
         return run
     ```
   - **Paragraph Spacing:** `space_before = Pt(1)`, `space_after = Pt(3)`, `line_spacing = 1.08`.
   - **Interoperability with Office Math (OMML):** When options contain equation fractions or radicals (e.g. `3/4`, `5/7`), append the `<m:oMath>` element inline before adding the tab run (`add_tab_run(p)`). Word smoothly advances to the next ruler tab stop immediately following inline math.

4. **Anti-Orphan Balanced Page Budgeting:**
   - Layout questions at exactly **10 questions per page** with balanced line spacing (`space_before = Pt(4)`, `space_after = Pt(1.5)`, `line_spacing = 1.12`) so that every page has an identical, clean, professional presentation.

5. **Solution & Explanations Section:**
   - Preceded by an explicit Page Break.
   - **Banner:** Full-width 1-cell table (`7.0 in`) with solid black background (`#000000`) and centered white bold text: `SOLUTION` (13 pt Cambria Math).
   - **Subheading:** Centered bold text `SECTION 01: [SUBJECT NAME]` (11 pt Cambria Math).
   - **Answer Key Grid:** 10-column bordered table with centered cell entries (e.g. `01. B`, `02. A`, etc.) in regular 10 pt Cambria Math. Cell widths: `0.70 in` each.
   - **Explanations:**
     - Bold heading `Explanations:` (11 pt Cambria Math).
     - Explanation question number & answer letter MUST BE REGULAR (UNBOLDED): e.g., `01. (B) ` must **not** be bold (`bold=False`).
     - Use OMML fractions, radicals, and auto-scaling bracket delimiters for all mathematical derivations and algebraic divisions.
     - Strictly enforce zero italics throughout.

---

## 6. Export Paths & Delivery

Always export generated documents to the user's primary project directory:
- **Default Export Folder:** `C:\AG\`
- **Files to Deliver:**
  1. `[Subject]_Questions_and_Solutions.docx`
  2. `[Subject]_Questions_and_Solutions.pdf` (converted via Word COM `convert_doc.ps1`)
