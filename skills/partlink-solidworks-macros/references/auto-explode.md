# auto explode

Macro: [Skycode - auto explode.bas](<../../../vba scripts/Skycode - auto explode.bas>)

## Purpose

Finds first-level assembly candidates among the actors visible in the active view, then applies a fixed linear explode to the descendant parts of each.

## Host and entry point

- Composer VBA through an Access host. Run `ExplodeFirstLevelSubAssembliesActiveView`.

## Inputs

- The active Composer document and its visible actors. The selection is not used. No prompts.
- Explode vector is +Z at a distance of 400.

## Outputs

- A `LinearExplode` call for each candidate's descendant parts.
- Phase diagnostics and a summary in the log.
- Log: `C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\auto explode.txt`.

## Side effects

- Changes the model's explode state. No reset is performed at the end.

## Gotchas

- The root ID is the first dot-delimited segment of the first usable actor path.
- Candidate assembly paths use the first two segments.
- By convention, assembly IDs start with `A` and part leaves with `P`.
- Path cleaning removes selection and modifiable wrappers, whitespace, and control characters.
- Uses early-bound Composer types. There is no VBA7 or Win64 handling.
