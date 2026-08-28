# Ledgerly Expense Tracker

Ledgerly is a vanilla HTML/CSS/JavaScript expense dashboard deployed with Netlify Functions. It uses two Excel workbooks stored in Netlify Blobs and server-side functions for account access, expense CRUD, and optional AI receipt reading. It does not use Java, Spring Boot, PostgreSQL, PHP, or a traditional application server.

## Run it

The application now requires the Netlify API for profiles, expenses, settings, and AI OCR. Do not open `index.html` directly with `file://`; that mode cannot run `/api/*` functions and will show a local-server message.

For local development, install dependencies and start the Netlify runtime:

```bash
npm install
npm run dev
```

Then open `http://localhost:8888`. Always use the **Local dev server ready** URL printed by Netlify. Do not open the separate `Static server listening` port, such as `3999` or `41097`; that port serves files only and returns `405 Method Not Allowed` for API requests. The CDN versions of SheetJS and Tesseract.js need an internet connection on first load.

### If localhost is refused in VS Code

The dev server runs inside the container. In VS Code, open the **Ports** panel, find port `8888`, and choose **Forward Port** or **Open in Browser**. Use the forwarded HTTPS URL that VS Code provides. Do not start `npm run dev` a second time while it is already running; `Could not acquire required 'port': '8888'` means an existing server already owns the port. Check it with:

```bash
curl http://localhost:8888/api/health
```

A working server returns `{"ok":true}`.

### Netlify deployment

Deploy the project directory to Netlify. The API stores the two workbooks in Netlify Blobs, not in the public site folder. Do not upload private Excel files as public static assets. The first registration creates the users workbook automatically; the first saved expense creates the user's worksheet in the expenses workbook. Enable Netlify Blobs for the site before using the application.

## Storage model

- The users workbook contains `Profiles`, `Settings`, and `ExchangeRates`. It stores usernames and password hashes, not raw passwords.
- Netlify Blobs holds the current workbooks server-side. Users never select the storage workbooks during normal use.
- The API creates and updates workbook sheets for CRUD operations. The UI reports API failures clearly.
- Profiles are local labels, not secure authentication. Anyone with the browser/workbook can access the data.

## Use

1. Register a username and password, or log in.
2. Add, edit, delete, search, filter, and sort expenses. LKR values are calculated from the stored original amount and exchange rate.
3. Manage categories and rates in **Categories & rates**. Settings are saved through the users workbook API.
4. Export a filtered worksheet, CSV, or print-ready report.

### Workbook layout

If you seed the Blobs manually, create these sheets and headers:

- `ledgerly-users.xlsx`: `Profiles` with `Username`, `Password Hash`, `Profile Name`, `Created Date`; `Settings` with `Profile`, `Categories`, `Default Currency`, `Theme Preference`; `ExchangeRates` with `Currency`, `Rate to LKR`, `Updated Date`.
- `ledgerly-expenses.xlsx`: one sheet per profile, using the profile name sanitized for Excel sheet rules. Each sheet has `ID`, `Date`, `Description`, `Category`, `Original Amount`, `Currency`, `Exchange Rate to LKR`, `Amount in LKR`, `Created At`, and `Updated At`.

Do not type a plain password into `Password Hash`. Register users through the application so the server creates a salted `scrypt` hash.

Receipt scanning is optional and review-first. The app sends the image to protected server vision AI when configured, asks specifically for purchased line items, quantity, product description, unit price, row total, currency, date, and category, and presents each item separately. If server AI is unavailable, it runs three local Tesseract preprocessing passes with English and Sinhala language data and applies line-item parsing that ignores tax, subtotal, payment, discount, and receipt metadata. Scan results remain in a scrollable review queue until the user clears them or scans again. Every suggestion must be reviewed before saving.

## Browser notes

The app uses SheetJS and Tesseract.js from CDNs for workbook compatibility and local fallback behavior. Currency rates are manual defaults and are never fetched from an API. Set `OCR_PROVIDER`, `OCR_MODEL`, and `OPENAI_API_KEY` in Netlify environment variables to enable server-side vision OCR. Passwords are salted and hashed with Node `scrypt`; raw passwords are never written to Excel.
