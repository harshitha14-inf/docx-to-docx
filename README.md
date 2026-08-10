# Docx Migration Tool

A simple explanation of what this tool does, how it works internally, and how to use it — written for someone opening this project for the very first time.

---

## 1. What problem does this tool solve?

Imagine you have a Word document (`.docx`) that was written a long time ago, in an old/inconsistent format. You now want its **content** (headings, paragraphs, tables, images, captions, formulas) placed inside a **new, official company template** — with the correct fonts, heading styles, table styles, and cover page — without you having to manually copy-paste everything and reformat it by hand.

This tool automates exactly that:

```
Old/messy .docx file  +  Official template  --->  New .docx file
   (your content)          (the target look)        (your content, in the new look)
```

It reads your original document, figures out what every piece of it *is* (is this a heading? a table? a picture? a formula? a caption?), and then rebuilds a brand-new document using the official template's styles, while keeping your original content, order, tables, and images intact.

At the end, it also **checks its own work** and tells you whether the conversion was successful (a "PASS/FAIL" report), so you don't have to manually compare 60-page documents yourself.

---

## 2. Who is this for?

Anyone who needs to convert an internal/legacy Word document into a standardized company template (e.g. a Datasheet template) without breaking formatting, losing images, or losing tables.

You do **not** need to know Python or any programming to use the finished tool — there is a simple point-and-click desktop app for that (see [Section 5](#5-easiest-way-to-use-it-the-desktop-app)). Developers/advanced users can also run it from the command line (see [Section 6](#6-using-it-from-the-command-line-advanced)).

---

## 3. How it works (the simple version)

Every conversion goes through 4 stages, always in this order:

```
 1. EXTRACT           2. BRAND/TRANSFORM        3. BUILD                4. VALIDATE
 ------------         --------------------      ------------            -------------
 Read the old        Apply small company-      Create the new         Compare old vs.
 .docx file and       specific adjustments      .docx file by          new file and
 turn it into a       to the extracted          combining your         report whether
 clean, internal      content (e.g. cover       content with the       anything was
 "model" of its       page fields, branding      OFFICIAL TEMPLATE's    lost or broken
 content              text)                      styles                 (PASS / FAIL)
```

### Stage 1 — Extract
The tool opens your original `.docx` (which is really just a `.zip` file full of XML files) and walks through every paragraph, table, image, and text box in order. For every single block of content, it decides what "kind" of content it is:

| What the tool sees in the file | What it decides it is |
|---|---|
| A short line styled "Heading 1/2/3/4" | **Heading** |
| A normal line of text | **Paragraph** |
| A grid of rows/columns | **Table** |
| A picture/drawing | **Image** |
| Text starting with "Figure 1" / "Table 2" | **Caption** |
| A line that looks like a math equation (e.g. `Z1 = R + jX`) | **Formula** (kept as a real, non-flattened equation) |
| Text inside a floating text box | **TextBox** |

This turns your entire document into a clean, structured list — completely independent of the original document's messy formatting.

### Stage 2 — Brand / Transform
A few small, company-specific adjustments are applied to that structured list (for example, preparing cover-page information like document title, revision number, or confidentiality marking). This step does **not** touch your actual body content — headings, paragraphs, tables, and images pass through unchanged.

### Stage 3 — Build
This is where the new file is actually created. The tool starts from the **official template** (its cover page, headers/footers, disclaimer page, fonts, and named styles like "Heading 1" or "TableCell") and inserts your content from Stage 1 into it — applying the template's actual style to each heading/paragraph/table/caption as it goes. The template controls 100% of the visual design; your original document only supplies the content and its order.

### Stage 4 — Validate
Once the new file exists, the tool re-opens BOTH the original and the new file and compares them on things like:
- Did every image survive?
- Did every table survive?
- Did every heading survive, with the correct style?
- Did every figure/table caption keep its numbering?
- Are there zero "forbidden" leftover placeholder phrases (like "Insert a bulleted list here") from the template?

It then prints a report ending in **PASS** or **FAIL**, so you immediately know if the output document is safe to use, or needs a second look.

---

## 4. Before you start (installation)

You need:
1. **Python 3.11+** installed on your computer.
2. This project folder, downloaded/cloned onto your machine.

Open a terminal inside the project folder and run:

```powershell
pip install -r requirements.txt
```

This installs everything the tool needs: `python-docx` (reads/writes Word files), `lxml` (reads the raw XML inside the Word file), and `PyQt6` (powers the desktop app).

---

## 5. Easiest way to use it: the Desktop App

If you just want to convert a document without touching any code, use the desktop app.

### Option A — You already have the `.exe`
If someone gave you `TechnicalDocumentConverter.exe` (found in the `dist/` folder after it's been built), just **double-click it**. No Python installation needed on your machine for this option.

### Option B — Run it from source
```powershell
python -m desktop_app.main
```

### Using the app
1. **Select Document Type** — choose what kind of document you're converting (currently, only **Datasheet** has a real template ready to use — see [Section 8](#8-supported-document-types)).
2. **Browse** and pick your input `.docx` file.
3. The app automatically suggests an output file name (`OUTPUT_<your file name>.docx`) in the same folder.
4. Click **Convert**.
5. Watch the progress bar — it shows each real stage of the process (Extracting → Branding → Building → Validating).
6. When it finishes, a success screen appears with buttons to **Open File** or **Open Folder** directly.
7. If something goes wrong, a clear error message is shown instead of a crash (e.g. "This file is currently open in Word — please close it and try again").

There's also a collapsible **"Show Log"** panel at the bottom if you want to see exactly what happened during the conversion (useful for troubleshooting).

---

## 6. Using it from the command line (advanced)

If you're comfortable with a terminal, you can run a conversion directly:

```powershell
python main_migration.py --source samples/input.docx --output samples/output.docx --document-type datasheet
```

What each option means:
- `--source` — path to your original `.docx` file.
- `--output` — where the newly generated `.docx` should be saved.
- `--document-type` — which template/rules to use (see [Section 8](#8-supported-document-types)). If you leave this out, the tool will simply ask you to type it in.

When it finishes, it prints a summary report in the terminal telling you whether the conversion **PASS**ed or **FAIL**ed, and why.

### Other useful scripts
```powershell
# Inspect a .docx file's structure/content without converting it (for debugging)
python scripts/analyser.py samples/file.docx

# Compare two .docx files against each other
python scripts/compare.py fileA.docx fileB.docx

# Run the full pipeline end-to-end on a known test file, as a sanity check
python tests/test_roundtrip.py
```

---

## 7. Understanding the validation report

After every conversion, you'll see a report like this:

```
========== VALIDATION REPORT ==========
CRITICAL CHECKS
----------------
[PASS] image_count                  Source=38 Output=38
[PASS] figure_count                 Source=19 Output=19
[PASS] caption_count                Source=41 Output=41
[PASS] heading_count                Source=94 Output=94
...
Overall Status : PASS
```

- **[PASS]** on a line means that check matched between your original file and the new file (e.g. the same number of images made it through).
- **[FAIL]** on a line means something didn't match — the report tells you the "Source" (original) value and the "Output" (new file) value side-by-side so you can see exactly what changed.
- **Overall Status** at the very end tells you the final verdict. If it says **FAIL**, open the new file and check the specific failed line(s) before sending it out.

There's also a second section, **Contract Compliance**, which checks softer rules like "did headings get the correct template style" and "is there zero leftover placeholder text from the template."

---

## 8. Supported document types

The tool is designed to support multiple document types, but each one needs its own **template file** before it can actually be used:

| Document type (`--document-type` value) | Template file it needs | Status today |
|---|---|---|
| `datasheet` | `templates/datasheet_template.docx` | ✅ Ready to use |
| `appnote` | `templates/appnote_template.docx` | ⏳ Not yet added |
| `specification` | `templates/specification_template.docx` | ⏳ Not yet added |
| `user_manual` | `templates/user_manual_template.docx` | ⏳ Not yet added |
| `release_note` | `templates/release_note_template.docx` | ⏳ Not yet added |

**To add support for a new document type:** simply place the correctly-named template `.docx` file into the `templates/` folder. The tool will automatically pick it up — no code changes needed. If a template file is missing, the tool will clearly tell you which file it expected and where.

---

## 9. The configuration files (`config/` folder) — what they control

You do not need to edit these to use the tool, but if you ever need to tweak its behavior, here's what each file does, in plain terms:

| File | What it controls |
|---|---|
| `document_type_rules.json` | The list of allowed document types, and which template file belongs to each one. |
| `style_map.json` | The "translation dictionary" between a generic content type (e.g. "this is a heading level 2") and the actual named style inside the template (e.g. `"Heading 2"`). Edit this if your template uses different style names. |
| `formatting_policy.json` | Rules about what to keep/strip during conversion — e.g. "keep bold text", "keep hyperlinks", "these exact phrases are template placeholder text and must never appear in a finished document" (see `forbidden_content`). |
| `acceptance_rules.json` | The pass/fail thresholds used by the validation report — e.g. "0 images are allowed to be lost" or "at least 95% of headings must match the expected style." |
| `cover_metadata.json` | Default text used to fill in the template's cover page fields (document type label, revision, confidentiality marking, etc.) when nothing more specific is provided. |

---

## 10. The templates folder (`templates/`)

- `datasheet_template.docx` — the actual blank company template used for datasheet conversions. This is a real Word file — open it in Word to see its cover page, disclaimer page, and named styles (Heading 1-4, TableCell, Bullet, Code, etc.).
- `style_inventory.json` — a reference list of every style name that exists inside the template, generated for convenience when editing `style_map.json`.
- `reference/` — extra original Infineon template files kept only as reference material; not used by the tool at runtime.

---

## 11. Project structure (what each folder/file is for)

```
docx-migration/
├── docx_migration/          <- the actual conversion engine (the "brain" of the tool)
│   ├── extractor.py         <- Stage 1: reads a .docx and classifies its content
│   ├── branding.py          <- Stage 2: applies company-specific adjustments
│   ├── builder.py           <- Stage 3: builds the new .docx using the template
│   ├── validator.py         <- Stage 4: compares old vs. new and reports PASS/FAIL
│   ├── models.py            <- the data shapes used to represent content internally
│   ├── document_types.py    <- the list of valid document types
│   ├── template_manager.py  <- finds/validates the correct template file to use
│   └── style_engine/        <- translates content types into real template style names
├── config/                  <- all the settings/rules described in Section 9
├── templates/                <- the official target templates described in Section 10
├── desktop_app/              <- the point-and-click desktop app (PyQt6)
│   ├── main.py               <- start the app from here
│   ├── ui/                   <- the visual windows/screens
│   └── services/             <- connects the app's buttons to the real conversion engine
├── scripts/                  <- small standalone helper tools for developers
├── tests/                    <- automated tests
├── samples/                  <- example input/output .docx files for trying things out
├── main_migration.py          <- command-line entry point (Section 6)
└── requirements.txt            <- the list of packages to install
```

---

## 12. Building your own desktop `.exe` (for distributing to non-technical users)

If you've made changes to the code and want to package a fresh, double-click-able `.exe` for someone else to use:

```powershell
python scripts/build_exe.py
```

This produces `dist/TechnicalDocumentConverter.exe` — a single file that runs the desktop app without needing Python installed on the receiving computer. It bundles the conversion engine, the config files, and the datasheet template inside itself.

**Important:** if you change any code in `docx_migration/` or `desktop_app/`, you must re-run this command to bake those changes into the `.exe` — the old `.exe` file will otherwise keep behaving exactly as it did when it was last built.

---

## 13. Frequently asked questions

**Q: My document type isn't in the list / I get a "Template file is missing" error.**
A: That document type doesn't have a template file yet. See [Section 8](#8-supported-document-types) — add the correctly-named `.docx` file into `templates/`.

**Q: The conversion failed with "file is open in another application".**
A: Close the document in Word (including any hidden background instances) and try again.

**Q: The report says FAIL — is the output file unusable?**
A: Not necessarily — a FAIL just flags a mismatch between old and new (e.g. one caption's style didn't match perfectly). Open the specific failed check in the new file and judge for yourself whether it matters for your use case.

**Q: Can this convert `.doc` (old Word 97-2003 format) or `.pdf` files?**
A: Not currently — the tool only reads real `.docx` files. If you have a `.doc` or `.pdf`, first "Save As" it as `.docx` in Word, then run it through this tool.

**Q: Where do I see detailed logs of what happened during a conversion?**
A: In the desktop app, click "Show Log". From the command line, the same information is printed directly to your terminal.

---

## 14. Current feature status

✅ Paragraph preservation
✅ Table preservation (including merged cells)
✅ Image preservation
✅ Heading preservation with correct template styles
✅ Figure/Table caption preservation (with live, auto-updating numbering)
✅ Formula/equation preservation (kept as real, editable equations — not flattened text)
✅ Cover page + header/footer field population
✅ Full validation report (PASS/FAIL)
✅ Desktop app (PyQt6) + standalone `.exe` packaging

⏳ Templates for `appnote` / `specification` / `user_manual` / `release_note` document types (structure is ready, template files are not yet supplied)
