# views [delete parts already in db]

Macro: [Skycode - views [delete parts already in db].bas](<../../../vba scripts/Skycode - views [delete parts already in db].bas>)

## Purpose

Deletes Composer views whose names appear in a list pasted by the user. Despite the filename, it does not query a database or Access.

## Host and entry point

- Composer VBA through an Access host. Run `DeleteExtraneousViews`.

## Inputs

- The active Composer document.
- A pasted list of view names. Names can be separated by newlines, commas, semicolons, tabs, or `#`.

## Outputs

- Deletes the matching views.
- Reports requested, matched, deleted, failed, and missing counts.
- Log: `C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\delete extraneous views.txt`.

## Side effects

- Deletion is not restored. Cancelling at either prompt prevents all deletion.

## Gotchas

- Matching is case-insensitive. Pasted names are trimmed and de-duplicated.
- `missing` is the unique requested names minus the unique names found.
- Review the pasted list and the preview counts before confirming.
