# set local pivot

Macro: [Skycode - set local pivot.bas](<../../../vba scripts/Skycode - set local pivot.bas>)

## Purpose

Sets each mesh part and assembly actor's pivot origin to its frame origin. Already-aligned actors are skipped.

## Host and entry point

- Composer VBA through an Access host. Run `SetPivotsLocalCoordSystem`.

## Inputs

- The active Composer document. It works on all mesh and assembly actors, not the current selection. No prompts.

## Outputs

- Rewritten `Actor.Location` for misaligned actors, written through `SetPropertySet`.
- Outcome and totals in the log.
- Log: `C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\set local pivot.txt`.

## Side effects

- Changes model location and pivot properties. Prior values are not restored.

## Gotchas

- The comparison tolerance is `0.0001`.
- The rebuilt XML keeps the axis vectors, sets `VolatilePivot="0"`, and sets `IsOk="1"` for Frame and Pivot.
- `CleanActorPath` strips a leading XML wrapper.
- Uses early-bound Composer types. There is no VBA7 or Win64 handling.
- Compare with [align-pivot](./align-pivot.md), which only changes the pivot origin.
