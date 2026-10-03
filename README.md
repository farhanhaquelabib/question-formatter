# Phoenix Image to Question Converter Skill

An autonomous agentic skill for converting images, screenshots, scans, and PDFs of questions into professionally formatted, publication-ready question papers and solution sheets in DOCX and PDF.

---

## 📌 Key Features & Rules

1. **Automatic Activation on Images/PDFs:**
   - Automatically activates whenever images or PDFs of questions are uploaded with requests to transcribe, create, format, or solve questions, **regardless of whether the skill name is explicitly mentioned**.

2. **Language Policy — English Only by Default:**
   - Questions are formatted in **English only** by default (`[Question Number]. [English Question Text]`).
   - Bangla translations are **omitted** unless the user explicitly prompts/asks for them.
   - If Bangla is explicitly requested, it is formatted unbolded in 9 pt Nirmala UI with standard English digits (`0–9`).

3. **Strict 0.75" Margins All Around:**
   - **Top Margin:** `0.75 in` (1080 dxa)
   - **Bottom Margin:** `0.75 in` (1080 dxa)
   - **Left Margin:** `0.75 in` (1080 dxa)
   - **Right Margin:** `0.75 in` (1080 dxa)
   - **Printable Width:** `7.0 in` ($8.5 - 0.75 - 0.75 = 7.0$ inches / 10,080 dxa).

4. **MCQ Options Layout — Never Use Tables:**
   - Plain text paragraph with OpenXML ruler tab stops (`<w:tabs>`):
     - **Option A:** `0.00 in` (`0 dxa`)
     - **Option B:** `1.75 in` (`2520 dxa`)
     - **Option C:** `3.50 in` (`5040 dxa`)
     - **Option D:** `5.25 in` (`7560 dxa`)

5. **Office Math (OMML) Equation Engine:**
   - **Vertical Stacked Fractions ($\frac{\text{num}}{\text{den}}$):** For all mathematical fractions and algebraic divisions (`6000/300`, `25/1.25`, etc.).
   - **Proper Equation Radicals:** Square roots ($\sqrt{x}$) with the vinculum extending over the entire radicand.
   - **Auto-Scaling Bracket Delimiters (`<m:d>`):** Parentheses automatically scale vertically to match the height of nested fractions.

6. **Strict Typography & Zero-Italics Policy:**
   - **Zero Italics:** Absolutely no italic words or math variables anywhere (`<m:nor/>` applied to all OMML math variables).
   - **English Questions:** Cambria Math 11 pt, **Bold**.
   - **Options:** Cambria Math 11 pt, **Regular (Unbolded)** with prefixes `A. `, `B. `, `C. `, `D. `.

7. **Balanced Page Budgeting:**
   - Exactly **10 questions per page** with balanced vertical spacing.

---

## 📁 Repository Structure

```text
phoenix-image-to-question-converter/
├── SKILL.md             # Core skill specification and autotrigger instructions
├── docx_helpers.py      # OpenXML styling, OMML math builders, and tab-stop helpers
├── convert_doc.ps1      # PowerShell Word COM DOCX -> PDF conversion script
├── scripts/
│   ├── docx_helpers.py  # Script copy of helpers
│   └── convert_doc.ps1  # Script copy of converter
└── README.md            # Skill overview and documentation
```

---

## 🚀 How to Install in Antigravity

Copy this folder into your global Antigravity skills directory:

```powershell
$dest = "$env:USERPROFILE\.gemini\config\skills\phoenix-image-to-question-converter"
New-Item -ItemType Directory -Path $dest -Force
Copy-Item -Path ".\*" -Destination $dest -Recurse -Force
```
