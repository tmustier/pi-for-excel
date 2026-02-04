# Pi for Excel — PoC (throw-away)

Scrappy proof-of-concept to validate platform capabilities. **Will be thrown away.**

## What we're testing

1. ✅/❌ Office.js add-in loads in Excel sidebar (macOS)
2. ✅/❌ pi-web-ui ChatPanel renders in the sidebar
3. ✅/❌ Office.js can read/write cells
4. ✅/❌ LLM API calls work from the webview (CORS)
5. ✅/❌ Pyodide (Python-in-WASM) loads and runs
6. ✅/❌ `Range.getDirectPrecedents()` works
7. ✅/❌ IndexedDB works in the webview
8. ✅/❌ `Worksheet.onChanged` events fire
9. ✅/❌ Basic Excel tool (read → LLM → write) end-to-end

## Setup

```bash
cd poc/
npm install

# Install trusted HTTPS certs (one-time)
mkcert -install
mkcert localhost

# Start dev server
npm run dev
```

## Sideload in Excel (macOS)

```bash
# Create the sideload directory if it doesn't exist
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/

# Copy manifest
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/

# Open Excel, go to Insert > My Add-ins > My Organization
# The add-in should appear. Click it to open the sidebar.
```

## Sideload in Excel (Windows)

Share the `manifest.xml` folder as a network share, or use:
```
\\localhost\path\to\poc\
```
Then in Excel: File > Options > Trust Center > Trusted Add-in Catalogs > add the path.
