# views [assemblies + top-level nonnumeric parts]

Macro: [Skycode - views [assemblies + top-level nonnumeric parts].bas](<../../../vba scripts/Skycode - views [assemblies + top-level nonnumeric parts].bas>)

## Purpose

Creates views for the root assembly, its subassemblies, and standalone parts. Part views are limited to parts that are not under a collected subassembly, and to parts with a blank or nonnumeric `Meta.PartNo`.

## Host and entry point

- Composer VBA through an Access host. Run `CreateAssemblyViews`, which connects to the active Composer document.

## Inputs

- The active Composer document and its actor paths and properties. No prompts, no database.

## Outputs

- Root, subassembly, and part views named as described in [views-assemblies-all-nonnumeric-parts](./views-assemblies-all-nonnumeric-parts.md).
- Log: `C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\views [assemblies + top-level nonnumeric parts].txt`.

## Side effects

- Changes visibility and zoom while creating views. Prior visibility and zoom are not restored.
- Releases Composer references and closes the log at cleanup.

## Gotchas

- Parts under a collected subassembly are excluded.
- Naming prefers `Meta.PARTNO_cfg`, then `Meta.PartNo_cfg`.
- Numeric part numbers are skipped.
