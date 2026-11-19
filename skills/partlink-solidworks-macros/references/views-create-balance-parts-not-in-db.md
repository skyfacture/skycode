# views [create balance parts not in db]

Macro: [Skycode - views [create balance parts not in db].bas](<../../../vba scripts/Skycode - views [create balance parts not in db].bas>)

## Purpose

Creates views only for the part numbers listed in a text file. Each requested part number is matched against the Composer actors, and a view is created for each match.

## Host and entry point

- Composer VBA through an Access host. Run `CreateViewsForPartsNotInDb`.

## Inputs

- The active Composer document.
- A `.txt` file with one part number per line, chosen in a file picker. If the picker is unavailable, the path can be typed in.
- Request normalisation: uppercase, with special characters removed except spaces and hyphens.

## Outputs

- Views named with the sanitised requested part numbers.
- The root view is renamed to a matched requested part number, if one matches.
- Match, create, and missing counts are shown at the end.
- Log: `C:\skycode\logs\views [create balance parts not in db].txt`.

## Side effects

- Hides all actors except the selected root or part, and zooms to fit. Prior visibility is not restored.
- Sets `CameraSmoothTime` to `0` before creating the root view and restores the previous value during cleanup, including cancellation and errors.
- Shows prompts for file selection, confirmation, and completion or error.

## Gotchas

- Matching prefers `Meta.PARTNO_cfg`, then `Meta.PartNo_cfg`, then `Meta.PartNo`.
- Assembly actors are preferred over part actors when both match.
- `STEP1_ONLY_TEST` is hardcoded to `False`.
- Root identification uses a fixed set of part numbers and IDs written into the macro.
