# meta partnos

Macro: [Skycode - meta partnos.bas](<../../../vba scripts/Skycode - meta partnos.bas>)

## Purpose

Sets each actor's `Meta.PartNo` to its configured part number. If no configured value exists, it sanitises the current value.

## Host and entry point

- Composer VBA through an Access host. Run `setMetaPartNoFromPartno_cfg`.

## Inputs

- The active Composer document and its mesh and view actors.
- `Meta.PartNo`, plus the configured `Meta.PARTNO_cfg` (preferred) or `Meta.PartNo_cfg`.

## Outputs

- Writes `Meta.PartNo` through the Composer property APIs.
- Log: `C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\meta partnos.txt`.

## Side effects

- Changes document properties. The log is overwritten. Nothing is saved.

## Gotchas

- Sanitising removes encoded numeric XML entities, decodes common XML entities, and keeps only ASCII letters, digits, space, dot, dash, and underscore.
- Three property-setting methods are tried in sequence.
- The collection of unique part numbers is declared but not used in the main loop.
