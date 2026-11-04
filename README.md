# Daily Report Sender

A static web page that turns the *Daily Performance Report* Excel workbook into WhatsApp-ready text, pictures and PDFs, then sends them from the user's own WhatsApp, in a fixed order.

Everything runs in the browser. The workbook is never uploaded anywhere.

## Using it

1. Open the site and give it the day's `.xlsx` file. You can:
   - pick it, drop it anywhere on the page, or paste a copied file;
   - open it with the installed app: on a computer, "Open with"; on a phone, share it to the app from WhatsApp or Files.
2. In about a second the page lists every output in send order, and fills each one in as it is made:
   1. the `SMS_To_IMC` text message;
   2. 23 report pictures, with no names attached, and a "." message between groups;
   3. 8 PDFs with the usual file names and page layout.
3. Send it:
   - **Send all to my WhatsApp** (computer, with the helper) posts everything, in order, to the chat with yourself in WhatsApp Web. That chat shows as "You", "Message yourself" or your own name. The steps tick off as they go.
   - **Share step 1 of 21** (phone, or no helper) opens the share screen for the next step at each tap. You pick your own chat each time.
   - **Download all (.zip)** saves every file at once.

Click a picture to view it full size. The arrow keys or a swipe move to the next one.

The site can be installed as an app (**Install app** in the header). It then opens without a network connection.

### The WhatsApp helper

A website cannot reach into WhatsApp Web by itself, so a small browser extension (`extension/`) does the posting. The site offers it as **Download the helper (.zip)**. To install it:

1. Open `chrome://extensions` (or `edge://extensions`).
2. Turn on Developer mode.
3. Click **Load unpacked** and choose the unzipped folder.

The helper works for whichever account is logged in to WhatsApp Web:

- it finds that account's own chat through WhatsApp Web itself, whatever the chat is called;
- it pastes the text, the pictures and the PDFs into it, and presses WhatsApp's send button each time;
- before every item it checks that the open chat is still the account's own.

Each send job is used at most once, so nothing is ever sent twice.

## What gets sent

Edit [`reports.js`](reports.js). Each entry is a named range from Excel's Name Box list and a format:

- `'text'`: the cells' text, one line per visible row;
- `'image'`: a PNG of the range;
- `'pdf'`: a PDF named `file`, laid out with `page`:
  - `paper` and `orientation`;
  - `margins` in points;
  - `fit` (`'width'`, or `'page'` to fit the whole page).

  Rows run on to more pages where a page is full, and at the sheet's own page breaks. A PDF with several names stacks them `perPage` to a page, like pictures pasted into a document.
- `'.'` in the picture list: a "." message, so WhatsApp shows the pictures before and after it as separate groups.

The PDF settings in `reports.js` were measured from the sample PDFs printed from Excel. The list order is the send order.

## Why it is fast

Excel takes minutes to open this workbook, because it recalculates every formula and loads all 27 sheets (one data sheet alone is 37 MB of XML). This page skips all of that:

- **Reads only what is needed.** The zip directory is read from the end of the file. Each part is then read straight from the File (`Blob.slice`) only when a report needs it, so 14 of the 27 sheets are never touched. Formula results come from the values Excel saved, so nothing is recalculated.
- **One engine per core.** Up to 8 Web Workers (`engine-worker.js`) each open the File themselves and turn reports into finished files: they parse, draw on `OffscreenCanvas`, encode PNG or JPEG and assemble the PDF. The page stays free and fills in each output as it arrives.
- **Ready before the file is.** The workers start, and load their fonts, while the user is still choosing a file.
- **Read once, shared.** The page reads the small shared parts once: sheet list, names, shared strings and styles (`book.js`, `readIndex`). It hands each worker a copy instead of every worker reading them again.
- **Lazy styles.** This workbook has 7,948 cell formats; each one is parsed only the first time a report uses it.
- **Scheduling** (`pipeline.js`). Reports go out biggest first. A worker prefers reports on sheets it has already parsed, so few sheets are parsed twice. If a worker dies, the others take over its reports.
- **No re-encoding for PDFs.** The PDF writer embeds the JPEG bytes as they are.

On a 4-core laptop a full run takes about 1.2–1.5 s from a cold start, with the first result after about 0.5 s. The old single-thread drawing took 1.8 s. The output is pixel-identical.

Browsers whose workers cannot draw (older Safari, for example) run the same engine on the page, one report at a time.

## Fidelity

The renderer reproduces Excel's own range picture:

- column widths in Excel's digit-width units, and row heights;
- fills, borders, fonts and rich text;
- number formats (via [SSF](https://github.com/SheetJS/ssf)), including Excel's General-format fitting and `####` overflow;
- merged cells, text overflow, and hidden rows and columns;
- conditional formatting, including Excel 2010 rules with cross-sheet references, top-N, and traffic-light, star and arrow icon sets;
- bar charts drawn from their cached values, and embedded pictures.

Built-in date formats and the `/` date separator follow an Indian-locale Excel.

On devices without Office fonts, text falls back to metric-compatible fonts that ship with the site (`vendor/fonts/`, SIL Open Font License): Carlito, Tinos, Arimo and Cousine.

## Development

No build step. Serve the folder and open it:

```sh
python -m http.server 8000
```

Self-check (formula engine, PDF writer, style tables, and, given a workbook, every configured report):

```sh
node test/check.mjs path/to/report.xlsx
```

Report data never belongs in this repository: `.gitignore` excludes `*.xlsx`.

Vendored: [fflate](https://github.com/101arrowz/fflate) (MIT), [SSF](https://github.com/SheetJS/ssf) (Apache-2.0) and the fonts above; see `vendor/`.
