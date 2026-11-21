# GST Rate Finder

A Streamlit app that tells you the **GST rate and related details for goods** in India, using an Excel
workbook of the CBIC rate schedules as its only source of rates. Nothing is guessed: if an item cannot be matched,
the app says so instead of inventing a rate.

> **This is a reference tool, not legal or tax advice.** Always confirm important rates on
> [cbic-gst.gov.in](https://cbic-gst.gov.in/gst-goods-services-rates.html) before using them on an invoice or return.

## What it does

| Page | Purpose |
|---|---|
| **Rate Finder** | Search by product name, common name or HSN code (full or partial). Shows GST rate, CGST / SGST / IGST, cess, taxability, conditions, legal reference, source and a verification status. Handles spelling mistakes and suggests corrections. |
| **GST Calculator** | Applies a rate you choose to an amount (tax-exclusive or tax-inclusive). |
| **Excel Bulk Processing** | Upload a sheet of product names and/or HSN codes (up to 1,000 rows) and download the results. Unresolved rows are kept and explained. |
| **Invoice Scan** | Take a photo of a printed invoice (or upload one from the gallery). The text is read locally (Tesseract OCR), you correct the item table, and each item is looked up in the workbook. The GST % printed on the invoice is only *compared* with the workbook - it never sets the rate. Details for the chosen item are shown on the same page. |
| **Search History / Favourites** | Stored locally, can be switched off or deleted. |
| **Data & Sources** | Shows the loaded workbook, rate slabs and official links; lets you replace the workbook. |

No login, no admin area, no telemetry, no external AI or OCR service. The layout adapts to phones and tablets
(a top menu replaces the sidebar on narrow screens).

## Repository contents

```
app.py                 The Streamlit application (single file)
app.ipynb              Google Colab notebook: installs everything, starts the app, prints a public link
requirements.txt       Python libraries (installed with pip)
packages.txt           System programs for Streamlit Community Cloud (Tesseract OCR for Invoice Scan)
.gitignore             Keeps private data (history, invoices, exports, notebook outputs) out of Git
data/UNIVERSAL_HSN_CODES_updated_3.xlsx   The HSN / GST rate workbook (see "Using the Excel file")
```

## Quick start - Google Colab (easiest, works from a phone)

1. Open `app.ipynb` in Google Colab (upload it, or open it from GitHub).
2. `Runtime > Run all`.
3. When asked, upload `UNIVERSAL_HSN_CODES_updated_3.xlsx`.
4. Open the `https://...trycloudflare.com` link printed by step 4. It also works in a phone browser.
5. When finished, run the **Stop** cell at the bottom.

## Quick start - Streamlit Community Cloud

1. Put `app.py`, `requirements.txt`, **`packages.txt`** and the workbook (`data/UNIVERSAL_HSN_CODES_updated_3.xlsx`) in a GitHub repository.
2. On share.streamlit.io create an app from that repository with `app.py` as the main file.
3. `packages.txt` makes the platform install the Tesseract OCR program. **Without this file Invoice Scan shows
   "Tesseract OCR program was not found"** - everything else still works. After adding it, use *Manage app > Reboot app*.
4. Cloud limits: storage is temporary (history/favourites reset on reboot), and an app set to *Public* can be opened by anyone with
   the link - set it to *Private* (invite viewers) if you will scan real invoices. Do not upload invoices with personal data
   to a public app.

## Quick start - your own computer

```bash
python -m venv .venv
source .venv/bin/activate            # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# Invoice Scan also needs the Tesseract program (not a Python library):
#   Ubuntu/Debian: sudo apt-get install tesseract-ocr tesseract-ocr-hin
#   macOS:         brew install tesseract
#   Windows:       install from https://github.com/UB-Mannheim/tesseract/wiki and add it to PATH

export GST_HSN_FILE="data/UNIVERSAL_HSN_CODES_updated_3.xlsx"   # Windows: set GST_HSN_FILE=...
export GST_DATA_DIR="gst_data"                                    # where history/favourites are kept
streamlit run app.py
```

Without Tesseract everything works except Invoice Scan, which shows what is missing.

---

## Using the Excel file (`UNIVERSAL_HSN_CODES_updated_3.xlsx`)

### What it is
The CBIC goods rate schedules (Schedules I-VII and the Nil-rate list) as one table, updated to **Notification
No. 9/2025-Central Tax (Rate), effective 22 Sep 2025**, plus later amendments listed in its *Update Notes* sheet
(including two goods entries effective 1 Nov 2026: retreaded tractor tyres at 5% and psyllium seeds at Nil).
Rates are stored as fractions (0.09 = 9%); the app converts them.

| Sheet | Used by the app? | Notes |
|---|---|---|
| `Table 1` | **Yes - the only sheet the app reads** | Goods rate schedules. Column I (*Remarks*) explains how each row was updated. |
| `Sheet1` | No | Empty. |
| `Update Notes` | No | Method, official sources and limits of the update. **Read it.** |
| `Services` | No (not yet) | Service rate entries, each tagged with a confidence level H / M / L. |
| `SAC_Code_Map` | No (not yet) | Maps 680 SAC codes to the nearest services entry, tagged H / M / L. |

### Precautions when using the Excel file

**Accuracy and currency**
1. **It is a snapshot, not the law.** It was prepared by a separate process (not by CBIC). Rates, exemptions and
   conditions change by notification. Check any rate you rely on against the CBIC notification text.
2. **Rows that kept old rates.** About 270 rows are marked *"Updated rates were not found - old rates retained"* in
   column I (headings shared by several rates, omitted and handicraft entries, malformed source cells). About 40
   entries still show **12% or 28%**, mostly handicraft rows (marked `*`). Treat any 12% / 28% result as **unverified**.
3. **Vehicles and split headings.** Some headings (for example 8703 cars, 8711, 8701) still list an old 28% line next to
   newer lines. The rate depends on engine size, length and other conditions - read the conditions, do not just take the
   highest or lowest rate.
4. **Compensation cess is Nil throughout.** The workbook states the cess was withdrawn (non-tobacco goods from
   22 Sep 2025; pan masala / tobacco from 1 Feb 2026). Confirm this for your product before invoicing.
5. **Value thresholds.** Some descriptions still quote the old Rs 1,000 apparel/footwear threshold; the new threshold is
   Rs 2,500 per piece/pair. The app shows the description text as written in the file.
6. **No effective date per row.** The app shows the effective date as "not stated". Use the *Update Notes* sheet.
7. **Services sheets are unverified.** Many rows are tagged **M** (medium) or **L** (low) confidence. Do not use
   service rates from them without checking the notification. The app does not read these sheets yet.
8. **Nov 2026 changes may be incomplete.** The notes cover retreaded tractor tyres and psyllium. News reports of the
   same 1 Nov 2026 changes also mention fertilisers and waste materials - verify those items yourself.

**Editing and replacing the file**
9. **Keep a copy of the original** before changing anything.
10. **Keep the layout.** `Table 1` must stay the **first sheet**, with the same 9 columns in the same order (schedule,
    S. No., chapter/heading/HSN, description, CGST, SGST, IGST, cess, remarks). Do not rename or reorder the columns,
    insert columns before column H, or merge cells in the data area. The app reads the first sheet that has at least 8
    columns and more than 20 rows.
11. **Do not delete rows to "clean up".** Omitted rows, `[Except ...]` text and `other than` wording are read by the
    app to apply exclusions. Changing them changes which goods match.
12. **Enter rates as fractions or percentages consistently** (0.18 or 18, never a mix in one file).
13. **After editing, reload.** On a computer, the app notices a changed file; on Colab, upload the new file in
    *Data & Sources* (it is validated before use) or re-run the notebook.
14. **Do not put personal or customer data in this workbook.** It is a rate reference only.

**For Excel bulk uploads (your own product lists)**
15. Use one row per product with a column for the name and/or HSN code. Maximum **1,000 rows**.
16. HSN codes should be text, not numbers - Excel drops leading zeros (`0402` becomes `402`). Format the column as Text.
17. Result files are written so that cells starting with `=`, `+`, `-` or `@` cannot run as formulas. Still, only open
    result files that you created yourself.
18. Do not upload sheets containing personal data (names, phone numbers, account numbers) - they are not needed.

---

## Precautions when using the software

**Check the result, always**
- A **Matched** result means the workbook has a single applicable entry. **Ambiguous** means several entries with
  different rates fit - the app deliberately does **not** pick one; read the conditions and choose yourself.
  **Unresolved / ITEM NOT FOUND** means nothing could be matched - check the spelling or use the HSN code.
- Fuzzy and common-name matching only *suggests* entries. The rate shown is always the one written in the workbook.
- An item's HSN code depends on what the goods really are, not only on their name. If you are unsure of the
  classification, confirm it with your tax professional.

**Invoice Scan**
- Works on **printed** invoices only. Handwriting is not supported. Tested only on generated sample invoices, so try it
  on your own invoices before relying on it.
- OCR makes mistakes (for example reading `I` as `l`, or a wrong digit in an HSN code). **Check every row in the table
  before pressing "Find GST".**
- Take a flat, sharp, well-lit, upright photo. On a phone, uploading from the camera app gives a sharper picture than the
  in-page camera. Small photos (under about 1,000 px wide) are read less accurately.
- The GST % printed on the invoice is **not** used to decide the rate. It is only compared with the workbook, and any
  difference is flagged for you to investigate.

**Privacy and data protection (India's DPDP Act)**
- Invoices contain business names, addresses, GSTINs and phone numbers - some belong to individuals. Upload only
  invoices you are entitled to process.
- The image and the text read from it are kept **only in your browser session**, are not saved, not added to history and
  not written to logs (logs keep counts only). Press **Clear invoice data** when you are done.
- On Google Colab the link is **public and has no login**, and traffic passes through Cloudflare. **Do not share the
  link**, do not run it for longer than needed, and stop the runtime afterwards. For regular business use, run the app on
  a private machine or server instead.
- Colab storage is temporary: history and favourites are lost when the session ends.
- History and favourites hold only search terms and rates. You can switch history off or delete it under *About &
  Privacy*.

**Keeping your Git repository clean**
- `.gitignore` already excludes the app's data folder, invoice images, exported result files and logs.
- **Clear notebook outputs before committing `app.ipynb`** (`Edit > Clear all outputs`) - the output contains your
  temporary public link.
- Do not commit real invoices, customer lists or result files, even to a private repository.

## Limits

- Goods only (HSN). Services (SAC) are not looked up yet.
- Bulk Excel: 1,000 rows. Invoice Scan: up to 5 images of up to 15 MB each (JPG, PNG, WEBP, HEIC, BMP, TIFF).
- Rates are as good as the workbook. When the workbook is out of date, the app is out of date.

## Troubleshooting

| Problem | Fix |
|---|---|
| "Reading text from images is not available" / "Tesseract OCR program was not found" | Tesseract is not installed. On Streamlit Cloud add `packages.txt` (see above) and reboot; on a computer install Tesseract; on Colab run the install cell. |
| "Workbook rejected" | The layout changed. Use the original file layout (see precaution 10). |
| Everything shows 12% / 28% warnings | You are using an older workbook. Use the `updated` file. |
| Colab link stops working | It changes on every start and stops when the runtime ends. Re-run the start cell. |
| Phone camera does not open | Use the https link (the Colab link is https) and allow camera access, or use *Upload / gallery*. |

## Sources

- CBIC goods rates: https://cbic-gst.gov.in/gst-goods-services-rates.html
- Notification No. 9/2025-Central Tax (Rate), dated 17-09-2025 and later amendments (see the workbook's *Update Notes*).
