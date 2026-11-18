# rename views in numerical order

Macro: [Skycode - rename views in numerical order.bas](<../../../vba scripts/Skycode - rename views in numerical order.bas>)

## Purpose

Renames the Composer views to `View 1`, `View 2`, and so on, in the order returned by the Composer API.

## Host and entry point

- Composer VBA through an Access host. Run `RenameViewsSequentially`.

## Inputs

- The active Composer document and its `GetAllViews` result. No prompts, no model or property inputs.

## Outputs

- Renamed views.
- Log: `C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\rename views in numerical order.txt`.

## Side effects

- Changes view names in the active document. The log is overwritten on each run. Nothing is saved or restored.

## Gotchas

- The order is the API's order, not a numeric sort of the existing names.
- Empty names are skipped, and the counter advances only for non-empty names.
