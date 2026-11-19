# kitify and bounding box

Macro: [Skycode - kitifyandboundingbox.bas](<../../../vba scripts/Skycode - kitifyandboundingbox.bas>)
Full specification: [kitify sync.md](<../../../docs/kitify sync.md#bounding-box-after-kitify>). Read it before changing this macro. Kitify behaviour is in [kitify](./kitify.md). The box calculation is in [sldworks bounding box](./sldworks-bounding-box.md).

## Purpose

Runs kitify to completion, then adds the bounding box to the generated combined root. The box is resolved only in `Default`. It is suppressed in every kit configuration. This is the canonical entry for a kitified root that will be imported into Composer.

## Host and entry points

- SOLIDWORKS VBA, or an Access host with the SOLIDWORKS type library (`sldworks.tlb`). The macro attaches to the running SOLIDWORKS session and never starts one.
- Import this module together with [Skycode - kitify.bas](<../../../vba scripts/Skycode - kitify.bas>) and [Skycode - sldworks bounding box.bas](<../../../vba scripts/Skycode - sldworks bounding box.bas>). Reimport changed modules before testing. A compiled isolated host does not update the user's Access project.
- The module name is `modKitifyAndBoundingBox`. It must not be `KitifyAndBoundingBox`: `Application.Run` cannot find a public procedure whose name matches its module.
- `KitifyAndBoundingBox`: no arguments. Uses the active saved assembly, lets kitify choose the next run folder, and prints the combined path to the Immediate window. It never shows a message box.
- `RunKitifyAndBoundingBox`: same run, and returns the combined assembly path.
- `AddBoundingBoxToKitifiedAssembly(combined)`: adds and confines the box on an already verified kitified root. Pass the combined assembly path. If omitted, the active document must be that assembly. Use this to finish a kitify output without running kitify again.

## Inputs

- A running SOLIDWORKS session and one saved production assembly. If several documents are open, activate the intended assembly before `KitifyAndBoundingBox`.
- The same PDM vault view and `12-KITS` folder required by kitify.
- For the box-only entry, a combined root whose parent folder is named `kitified-generated_<source base>_<yyyymmdd>_<run>`, beside the unchanged source assembly.

## Outputs

- The kitify run folder beside the source, `kitified-generated_<partno>_<yyyymmdd>_<run>`. The run number is the next free integer for that part number and date, starting at 1. It does not continue the previous day's number.
- The combined root from kitify, reopened and saved after the box is added. The source file is not saved.
- A proxy part `<combined>_BoundingBox.SLDPRT` and the bounding-box text report beside the combined root.
- `kitifyandboundingbox.txt` in the run folder. It records the combined path, the source size and time, and the bounding-box report.
- Failure text, when the wrapper fails, at `C:\skycode\logs\kitifyandboundingbox.txt`.
- PDM No responses queued by this wrapper at `C:\skycode\logs\kitifyandboundingbox-declined.txt`. Kitify may also write `declined-checkouts.txt` in the run folder.

## What it does

1. Records the source file size and last-write time.
2. Removes the Access OLE message filter for the out-of-process host, and starts the PDM No helper.
3. Calls `KitifyDrifter("", "", False)` and waits until it returns. Kitify closes the generated assembly and reactivates the source.
4. Checks that the returned assembly exists inside the source folder and is not the source file. A same-parent check is wrong: the assembly lives in the dated run folder.
5. Reopens that assembly with `OpenDoc6`, activates `Default`, and calls `GetCurrentModelBoundingBox(False)`.
6. Suppresses the envelope component and `Skycode_BoundingBox` in every non-Default configuration, then checks that `Default` still has both resolved.
7. Saves only the generated assembly with `Save3(1)`, checks the source file again, and leaves the generated assembly active on `Default`.

## Side effects and restoration

- Kitify's own prompt changes and document budget still apply. This wrapper also toggles external-reference save prompts, update-error warnings, referenced-document warnings, and auto-dismiss of open messages, then restores the values it changed.
- The wrapper starts its own hidden PDM helper for the post-kitify open and save. It requests No only on a visible `#32770` dialog titled `SOLIDWORKS PDM`, owned by the SOLIDWORKS process, with exactly one No button. `QUEUED NO` means the command was posted, not that the dialog accepted it. On 2026-11-19 this helper queued No for the checkout dialog on `BG 006 019 67.SLDPRT` during kitify and again during the box reopen. The run continued.
- Does not save or clear the dirty flag of the source assembly. Inspection can leave it dirty in memory.
- Closes unsaved documents and a leftover `box-config-probe-assembly.SLDASM` before kitify. It does not close other saved models. Kitify still refuses to start when another saved visible document is open.

## Gotchas

- Do not add the box before kit configurations exist. `AddConfiguration3` copies the active configuration, so a resolved box in `Default` is copied, resolved, into every kit configuration.
- `InsertEnvelope` is the working insert path. `AddComponent4` and `AddComponent5` returned `Nothing` on the 2026-11-19 probe. `IsEnvelope` is read-only. Inserting the envelope after the kit configurations exist still resolves it in every existing configuration.
- Do not delete the envelope from a kit configuration. `DeleteSelection2` also removes it from `Default`.
- Confinement is suppression. `SetSuppression2(0)` applies to the active configuration and returns `2` on success, not `0`. `ShowConfiguration2` returns False when that configuration is already active; treat that as success.
- On the 2026-11-19 Composer import of `kitified-generated_7397 0500 28_20261119_1`, Composer did not see the suppressed box in the kit configurations. The box stayed resolved in `Default`. Enable **Import SOLIDWORKS assembly envelope**.
- Normalize SOLIDWORKS Booleans with `CBool(flag <> False)`.
- A VBA error dialog ended with End returns `0x800A9C68` (`OperationStopped`) to an out-of-process `Application.Run`. Read the dialog text before ending it. The wrapper also writes that failure to `C:\skycode\logs\kitifyandboundingbox.txt`.

## Canonical live result (2026-11-19)

`KitifyAndBoundingBox` completed against `C:\skycode\partlink cad sandbox\7397 0500 28.SLDASM`.

- Output: [kitified-generated_7397 0500 28_20261119_1](<../../../partlink cad sandbox/kitified-generated_7397 0500 28_20261119_1/kitifyandboundingbox.txt>). The previous day's highest folder was `_20261118_7`, so this run was `_1`, not `_8`.
- The source file stayed 13,789,862 bytes, last written 18/11/2026 8:40:47 PM.
- Kitify verified `Default` plus 16 kit configurations. The box was added afterwards: 517 x 271 x 190 mm, resolved only in `Default`.
- The user imported that root into Composer and confirmed that Composer does not see the box in the kit configurations.
