# Referal-Conversion

Convert **Student Assessment Results by Class Code** PDFs into one Excel workbook — one tab per PDF — laid out exactly like the existing results workbooks (navy header row, thin borders, failed rows shaded, PDF markers such as `N` kept as cell comments).

Everything runs locally. No internet connection or account is needed.

## For office users (no Python needed)

Download from the repository's **Releases** page (or the latest **Actions → Build apps** run → *Artifacts*):

| Computer | File | How to start |
|---|---|---|
| Windows | `Assessment PDF to Excel.exe` | Double-click. First time, Windows may show *"Windows protected your PC"* → **More info** → **Run anyway**. |
| Mac (Apple Silicon: M1/M2/M3/M4) | `Assessment-PDF-to-Excel-mac.dmg` | Open the .dmg, drag the app onto **Applications**. First time: right-click the app → **Open** → **Open** (or System Settings → Privacy & Security → **Open Anyway**). |

The warnings appear because the app is not signed with a paid Microsoft/Apple developer certificate. Then follow the steps under [Desktop app](#desktop-app), or click **Help** inside the app.

## Building the .exe / .dmg

Automatic: every push to `main` builds both on GitHub (`.github/workflows/build.yml`). To publish a release with both files attached:

```sh
git tag v1.0.0
git push origin v1.0.0
```

Manual (each must be built on its own operating system):

```sh
bash installer/build_mac.sh          # macOS -> dist/Assessment-PDF-to-Excel-mac.dmg
installer\build_windows.bat          # Windows -> dist\Assessment PDF to Excel.exe
```

The packaged app also works from a terminal with the same options as `convert.py` (e.g. `"Assessment PDF to Excel.exe" FOLDER -o out.xlsx`).

## Setup (running from source)

Python 3.10+ with Tkinter (included in the standard python.org installers).

```sh
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
python -m pip install -r requirements.txt
```

## Desktop app

```sh
python convert.py
```

1. **Add PDFs…** — select one or more PDFs (Ctrl-click / Command-click for several).
2. Tab names default to the PDF filename. Double-click a row to rename a tab.
3. Enter the **workbook filename** and **save folder** (defaults to Documents).
4. Optional: enter an **index sheet name** to add a first sheet listing every PDF. Leave blank for the same layout as the existing workbooks.
5. Click **Convert to Excel**.

The **Help** button in the app repeats these steps.

## Command line

```sh
# one folder of PDFs -> one workbook
python convert.py NPTMF2503MBA --output NPTMF2503MBA.xlsx

# custom tab names (one --tab per PDF, in order; --tab '' keeps the PDF name)
python convert.py a.pdf b.pdf -o results.xlsx --tab "Finance" --tab ""

# options
#   --recursive          include PDFs in subfolders
#   --main-sheet NAME    add a first index sheet
#   --overwrite          allow replacing an existing workbook
```

## Output rules

- Columns: S.No, Name, Student ID, Course, each assessment (with %), Total Marks, Grade Point, Grade, Remarks.
- Marks, totals and grades are copied as printed — never recalculated. Student IDs stay text; marks stay numbers.
- Rows with grade F, F+, F- or U are shaded light red. PDF markers (N, EC, NQ, A, …) become comments on the mark cell.
- Excel tab names: max 31 characters, `\ / * ? : [ ]` replaced with `_`, duplicates get ` (2)`, ` (3)`.
- Every PDF is validated before anything is written: student counts are reconciled per page; duplicate IDs, broken serial numbers, unreadable marks or weights not totalling 100% stop the conversion. An existing workbook is never left half-written.
- Only this report format is supported. Scanned (image-only) or encrypted PDFs are rejected.

## Tests

```sh
python -m unittest discover -s tests -v
python verify_samples.py
```

`verify_samples.py` compares extraction against reference workbooks. It looks for sample folders laid out as `FOLDER/FOLDER.xlsx` next to that class's PDFs. Sample data contains student records and is **not** committed (see `.gitignore`); tests that need it are skipped when it is absent.

## Files

| File | Purpose |
|---|---|
| `app.py` | Entry point for the packaged app |
| `installer/` | PyInstaller recipe and build scripts |
| `convert.py` | CLI, batch validation, Excel export |
| `desktop.py` | Desktop window (Tkinter) |
| `pdf_extractor.py` | Strict PDF parsing |
| `verify_samples.py` | Audit against reference workbooks |
| `tests/` | Unit tests |
