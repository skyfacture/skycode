# composer connect diagnostic

Macro: [Skycode - composer connect diagnostic.bas](<../../../vba scripts/Skycode - composer connect diagnostic.bas>)

## Purpose

Diagnoses Composer COM registration and connection. It reports whether the macro can attach to a running Composer instance, and whether it can launch one.

## Host and entry point

- Run `DiagComposer` from the Access VBA Immediate window or as a macro.

## Inputs

- None. It probes the candidate ProgIDs and the active document if one is open. No prompts.

## Outputs

- Registry and connection probes, written to the log and the Immediate window.
- A completion message box.
- Log: `C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\composer diagnostic.txt`.

## Side effects

- Does not change the model.
- If attaching fails, it may launch a new Composer instance. It does not shut Composer down.

## Gotchas

- `AttachComposer` tries `DSComposer.Application`, then `Composer.Application`. Launch uses the same two IDs.
- `DumpProgId` also checks `ComposerPlayer.Application`.
- VBA7 and Win64 are reported with conditional compilation.
- Use this macro first when any Composer macro fails to connect.
