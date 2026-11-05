# Shrishti Enterprises Chart-to-Bill

Mobile-friendly local pilot for reviewing chart quantities and producing sample Credit Bill or Delivery Challan workbooks.

## Run locally

Requirements: Node.js 24 and npm. For image extraction, install Ollama and pull a vision model supported by Ollama:

```powershell
npm install
Copy-Item .env.example .env
ollama pull qwen3-vl:4b
```

Start the API and client in separate terminals:

```powershell
npm run dev:api
npm run dev
```

Open the Vite URL printed in the client terminal. Without Ollama, use **Load example data** to exercise review, configuration, and Excel generation. To choose another local model or Ollama endpoint, edit `.env`.

## Pilot scope

- Accepts JPG, PNG, and WebP chart images up to 12 MB; signatures are checked server-side.
- Runs extraction through local Ollama and asks the model to return structured rows, leaving unclear values empty for human correction.
- Allows up to 4,096 generated tokens for chart extraction; split very large charts into smaller images if extraction still reports truncated JSON.
- Extracts the chart's scheme/product heading as editable source text, preserving Kannada Unicode; the General/SCP/TSP allocation column remains a separate classification for template selection.
- Lets the invoicing person set an independent starting number for Credit Bills and Delivery Challans; the form advances the number after each successful batch.
- Supports editable ADA, RSK, scheme, crop, and quantity rows; Delivery Challans require an RSK.
- Fills the supplied `Credit bill Format.xlsm` and `DC Format.xlsm` layouts, preserving their merged cells, formatting, formulas, and page setup.
- Uses Decimal arithmetic for GST and replaces the template's unavailable `SpellNumber()` formula with generated amount-in-words text.
- Supports inclusive and exclusive GST settings and adds amount-in-words text.
- Downloads generated files as macro-free `.xlsx`; neither uploaded template contains a VBA project.

Chart quantities currently map to the template's first goods line (`G13`); product descriptions and HSN remain as supplied in the templates. Multiple chart rows for the same ADA/scheme bill are blocked until an explicit crop-to-line-item map is configured.

This is not production billing software. Jobs, bill history, and the separate Credit Bill/Delivery Challan number sequences are in memory and reset when the API restarts. There is no authentication, database, durable storage, PDF export, audit log, or backup. Confirm the chart-to-product mapping and accountant-approved tax settings before issuing generated files as official bills.

## Checks

```powershell
npm test
npm run lint
npm run build
```
