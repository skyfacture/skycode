# views [assemblies + all nonnumeric parts]

Macro: [Skycode - views [assemblies + all nonnumeric parts].bas](<../../../vba scripts/Skycode - views [assemblies + all nonnumeric parts].bas>)

## Purpose

Creates views for the root and each unique subassembly, then for every part with a blank or nonnumeric `Meta.PartNo`, including parts under subassemblies.

## Host and entry point

- Composer VBA through an Access host. Run `CreateAssemblyViews`.

## Inputs

- The active Composer document and its actor metadata. No prompts, no user-supplied paths, no database.

## Outputs

- Root view named from its part number, or `Root_<ID>`.
- Subassembly views named from their part number, or `SubAssy_<path>`.
- Part views named from `Meta.PARTNO_cfg`, then `Meta.PartNo_cfg`, then `Meta.PartNo`, with `Part_<path>` as the fallback.
- Log: `C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\views [assys + all nonnumeric parts].txt`.

## Side effects

- Changes visibility so that only the selected assembly or part shows, and zooms to fit. Prior visibility is not restored.
- Does not save or restore camera smoothing or collaborations.

## Gotchas

- Numeric means every character passes `IsNumeric`. Blank values count as nonnumeric.
- Parts are detected by a final actor-path segment that begins with `P`.
- Views are deduplicated by name. A duplicate name is skipped.
