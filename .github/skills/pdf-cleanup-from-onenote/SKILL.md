---
name: pdf-cleanup-from-onenote
description: 'Convert OneNote MHT/web-archive exports to a readable A4 PDF. Use when an exported OneNote document is clipped, oversized, or needs one form page per PDF page.'
---

# Clean up a OneNote PDF

Convert the MHT locally with Microsoft Word. Do not send the document to online conversion services; OneNote exports may contain confidential information.

## Procedure

1. Keep the original MHT unchanged. Confirm Microsoft Word is installed.
2. Run [convert.ps1](./scripts/convert.ps1) with the MHT path and, optionally, a PDF output path:

   ```powershell
   .\convert.ps1 -InputPath 'C:\path\to\export.mht'
   .\convert.ps1 -InputPath 'C:\path\to\export.mht' -OutputPath 'C:\path\to\clean.pdf'
   ```

3. The script opens the MHT read-only in Word, sets portrait A4 with 18-point margins, and scales oversized inline content proportionally to fit the printable width and height. Word's headless browser mode may not render MHT reliably.
4. Check the resulting PDF's page count and inspect representative pages, especially the first and last pages and any page that contains tables or unusually large images. Ensure text and form fields are not clipped and that each intended form page is on its own PDF page. Scaling can change pagination; do not assume the MHT's initial Word page count is the intended final count or insert page breaks blindly.
5. The script verifies the PDF signature, A4 page boxes, and page count before replacing the requested output. If verification fails, it reports an error and preserves the existing output.
