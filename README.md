# Notice Generator CLI

Generate bank notices from an Excel/CSV accounts list using a Word template, with smart styling that mirrors your template’s fonts, spacing, and layout.

## Features

- Template-driven styling: derives font, size, spacing, column widths from `sample_notice.docx`
- Nodal officer block: renders the bank name below “NODAL OFFICER” in the exact template style
- Table fill: clears old rows and inserts compact, bordered rows with account number, holder, IFSC
- IFSC-aware: validates format (`XXXX0YYYYYY`) and maps first 4 characters to bank name
- Flexible input: reads `.xlsx` or `.csv` and auto-detects common header variants
- Tone option: optional emphasis for urgent/friendly documents without altering template fonts

## Requirements

- Python 3.9+
- Packages: `pandas`, `openpyxl`, `python-docx`
- Quick setup helper:
  - `python requirement.py --write` to generate `requirements.txt`
  - `python requirement.py --install` to install any missing packages

## Template Expectations

- Bank-name placeholder text (default: `ICICI BANK`) present somewhere in the document
- A 3-column accounts table with a header that contains `account` and `ifsc`
- The paragraph immediately following the text “NODAL OFFICER” has the style you want for the bank name

## Data Expectations

- Excel/CSV with three columns (names are detected flexibly):
  - Account number: includes `account` and `number`, or variants like `Account No`, `A/C No`, `Acc No`
  - Account holder: includes `account` and `name`, or variants like `Beneficiary Name`, `Account Holder`
  - IFSC: includes `ifsc`

## CLI

```
python makenotice.py <accounts.xlsx|csv> <template.docx> [options]
```

Options:

- `-o, --output-dir <dir>` Output directory (default: `notices_output`)
- `--placeholder <TEXT>` Placeholder in template to replace with bank name (default: `ICICI BANK`)
- `--tone <formal|urgent|friendly|auto>` Optional tone (default: `formal`)
- `--font-name <NAME>` Fallback font if template lacks font (default: `Bookman Old Style`)
- `--font-size <SIZE>` Fallback size in points (default: `8`)

Helper script:

- `python requirement.py --install` installs missing dependencies
- `python requirement.py --upgrade` upgrades dependencies to latest

## Quick Start (Windows)

- Example with template-driven styling:
```
python makenotice.py "C:\Users\wolf\Desktop\developer things\makenotice\fake_bank_accounts.xlsx" \
  "C:\Users\wolf\Desktop\developer things\makenotice\sample_notice.docx" \
  -o notices_template_driven --placeholder "ICICI BANK" --tone auto
```

- Example with CSV:
```
python makenotice.py "C:\Users\wolf\Desktop\developer things\makenotice\counterparty_accounts.csv" \
  "C:\Users\wolf\Desktop\developer things\makenotice\sample_notice.docx" \
  -o notices_from_csv
```

- Explicit font override (if needed):
```
python makenotice.py accounts.xlsx template.docx -o notices --font-name "Bookman Old Style" --font-size 8
```

## How It Works (Summary)

- Reads data, validates IFSC, groups records by IFSC
- Resolves bank name from IFSC prefix or falls back to `<CODE> BANK`
- Loads the template, extracts baseline font/spacing and header widths
- Replaces all placeholder occurrences with the bank name
- Updates the accounts table header and data rows with borders, column widths, and compact spacing
- Places the bank name below “NODAL OFFICER” using the template’s next-paragraph style
- Saves one `.docx` per IFSC group into the output directory

Output naming:

- `Notice_<BANK_NAME>_<IFSC>.docx` written to the chosen output directory

## Troubleshooting

- Columns not detected: ensure headers contain the keywords described above; check console output for detected columns
- Rows without borders: confirm the template’s 3-column accounts table is present; the tool enforces borders per cell and table
- Bank name style differs: verify the paragraph immediately after “NODAL OFFICER” is styled as desired
- Cannot delete output files: close any open Word windows using those files, then delete the folder

## Safety & Notes

- No secrets are logged or written
- Generated files inherit as much of the template styling as possible; explicit CLI font options are only fallbacks
- Unknown IFSC bank codes render as `<CODE> BANK` to avoid failures
