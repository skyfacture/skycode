# align pivot

Macro: [Skycode - align pivot.bas](<../../../vba scripts/Skycode - align pivot.bas>)

## Purpose

Moves the pivot origin of every mesh part and assembly actor to its frame origin. Actors already aligned within tolerance are skipped.

## Host and entry point

- Composer VBA through an Access host. Run `AlignPivotToFrame`.

## Inputs

- The active Composer document. It works on all mesh and assembly actors, not the current selection. No prompts.

## Outputs

- Modified `Actor.Location` values, written through `SetPropertyMap`.
- Per-actor results and totals, written to the log and the Immediate window.
- Log: `C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\align pivot.txt`.

## Side effects

- Changes model pivots. No rollback. The log is closed and references released at cleanup.

## Gotchas

- The comparison tolerance is `0.0001`.
- Frame and pivot axis vectors are preserved, with defaults if they are missing.
- `CleanActorPath` removes a leading XML prefix and trims the path.
- Uses early-bound `DSComposer` types. There is no VBA7 or Win64 handling.
- Compare with [set-local-pivot](./set-local-pivot.md), which does the same job by rewriting the whole location.
