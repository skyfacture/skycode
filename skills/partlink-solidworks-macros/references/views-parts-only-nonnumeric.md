# views [parts-only nonnumeric]

Macro: [Skycode - views [parts-only nonnumeric].bas](<../../../vba scripts/Skycode - views [parts-only nonnumeric].bas>)

## Purpose

Creates a view for each part with a blank or nonnumeric `Meta.PartNo`. Assemblies get no views.

## Host and entry point

- Composer VBA through an Access host. Run `CreateAssemblyViews` with the target Composer document active.

## Inputs

- The active Composer document and its part actors. No prompts, no database.

## Outputs

- One view per qualifying part.
- Log: `C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\views [parts-only nonnumeric].txt`.

## Side effects

- Changes visibility and zoom to isolate each part. Prior state is not restored.
- Closes the log and releases object references at cleanup.

## Gotchas

- Numeric means every character passes `IsNumeric`. Blank values count as nonnumeric.
- Naming preference is `Meta.PARTNO_cfg`, then `Meta.PartNo_cfg`, then `Meta.PartNo`. The fallback is `Part_` plus the actor path.
- Duplicate view names are skipped.
