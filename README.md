# Daily Report Sender

A static web page that turns the *Daily Performance Report* Excel workbook into WhatsApp-ready text, images and PDFs, then shares them from the user's own WhatsApp, in a fixed order.

Everything runs in the browser. The workbook is never uploaded anywhere.

## Using it

1. Open the site and pick the day's `.xlsx` file.
2. In about 2 seconds the page lists the outputs in send order, in three steps:
   1. the `SMS_To_IMC` text message;
   2. 23 report images;
   3. 8 PDFs, including one combined PDF of the four CCL coal / OBR / offtake / rake reports.
3. Press **Send to WhatsApp** on each step and pick the contact or group. On a phone this opens WhatsApp directly. On Windows it works through the share dialog when the WhatsApp desktop app is installed.

Where the browser cannot share files, use **Download** and attach the files in WhatsApp.

## What gets sent

Edit [`reports.js`](reports.js). Each entry is a named range from Excel's Name Box list and a format:

- `'text'`: the cells' text, one line per visible row;
- `'image'`: a PNG of the range;
- `'pdf'`: a PDF. Give a list of names to get one PDF with a page per range.

The list order is the send order. Consecutive entries of the same format are shared together.

## Why it is fast

Excel takes minutes to open this workbook, because it recalculates every formula and loads all 27 sheets (one data sheet alone is 37 MB of XML). This page skips all of that:

- **Reads only what is needed.** It reads the zip directory itself and inflates only the 12 or 13 sheets the reports live on. Formula results come from the values Excel saved, so nothing is recalculated.
- **Parallel parsing.** Each needed sheet's compressed bytes go to a pool of Web Workers. The biggest sheet goes first, onto the least-busy worker. Each worker inflates and parses its sheet with a DOM-free scanner. Meanwhile the main thread reads styles, theme, charts and pictures.
- **Pipelined drawing and encoding.** Reports are drawn one at a time on `OffscreenCanvas`, biggest first. Each finished drawing moves without copying (`transferToImageBitmap`) to a pool of encoder workers. They produce PNG or JPEG in parallel while the next report is drawn. Only a few canvases exist at once, so phones do not run out of memory.
- **No re-encoding for PDFs.** The PDF writer embeds the JPEG bytes as they are.

The page shows its timing after every run.

## Fidelity

The renderer reproduces Excel's own range picture:

- column widths in Excel's digit-width units, and row heights;
- fills, borders, fonts and rich text;
- number formats (via [SSF](https://github.com/SheetJS/ssf)), including Excel's General-format fitting and `####` overflow;
- merged cells, text overflow, and hidden rows and columns;
- conditional formatting, including Excel 2010 rules with cross-sheet references, top-N, and traffic-light, star and arrow icon sets;
- bar charts drawn from their cached values, and embedded pictures.

Built-in date formats and the `/` date separator follow an Indian-locale Excel. Fonts fall back to metric-compatible web fonts (Carlito, Tinos, Arimo, Cousine) on devices without Office fonts.

## Development

No build step. Serve the folder and open it:

```sh
python -m http.server 8000
```

Self-check (formula engine, PDF writer, and, given a workbook, every configured report):

```sh
node test/check.mjs path/to/report.xlsx
```

Report data never belongs in this repository: `.gitignore` excludes `*.xlsx`.

Vendored: [fflate](https://github.com/101arrowz/fflate) (MIT) and [SSF](https://github.com/SheetJS/ssf) (Apache-2.0); see `vendor/`.
