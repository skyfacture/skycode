# Identify Root

Macro: [Skycode - Identify Root.bas](<../../../vba scripts/Skycode - Identify Root.bas>)

## Purpose

Finds the top-level assembly IDs whose part numbers match a fixed list, selects and isolates them, and creates and names a view for each.

## Host and entry point

- Composer VBA through an Access host. Run `IdentifyRoot`.

## Inputs

- The active Composer document. Actors are taken from `GetAllActors(1)`. No prompts.
- A hardcoded list of four part numbers, plus a list of assembly IDs, both written into the macro.

## Outputs

- Selected matching assemblies, with visibility changed, zoom applied, and views created and renamed.
- Progress and errors in the log.
- Log: `C:\Temp\ComposerDebug.txt`. This path is hardcoded and is not the usual logs folder.

## Side effects

- Changes selection, visibility, zoom, and the document's views. Visibility and selection are not restored, and created views are not removed.

## Gotchas

- The top-level ID is the text before the first period.
- `StripMetaPartNo` extracts a `Value` attribute, or strips XML tags.
- The part-number and ID lists must be updated in the macro when the root set changes.
- Uses early-bound Composer types. There is no VBA7 or Win64 handling.
