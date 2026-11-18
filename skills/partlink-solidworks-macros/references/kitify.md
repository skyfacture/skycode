# kitify

Macro: [Skycode - kitify.bas](<../../../vba scripts/Skycode - kitify.bas>)
Full specification: [kitify sync.md](<../../../docs/kitify sync.md>). Read it before changing this macro. This reference summarises it and does not replace it.

## Purpose

Builds one verified kitified assembly from the active production SOLIDWORKS assembly. The output is one root with `Default` plus one configuration per PDM kit assembly in the sibling `12-KITS` folder. Kit membership comes from the PDM kit assemblies, not from the source file.

## Host and entry points

- Runs in SOLIDWORKS VBA, or in Access with the SOLIDWORKS type library (`sldworks.tlb`) referenced. Access hosting is supported; the macro removes the OLE message filter during the run.
- Import [Skycode - sldworks bounding box.bas](<../../../vba scripts/Skycode - sldworks bounding box.bas>) into the same VBA project: Kitify calls `GetCurrentModelBoundingBox(False)`. Reimport changed modules before testing; a compiled isolated test host does not update the user's Access project.
- `SkycodeKitify`: runs with defaults. Writes the result to the Immediate window and to `run.txt`. It never shows a message box.
- `KitifyDrifter(sourcePath, outputFolder, showMessage)`: all arguments are optional. An explicit `sourcePath` must already be open. An explicit `outputFolder` must not exist.
- `TestKitifyMetadata`: regression checks for metadata and provenance. Prints PASS.

## Inputs

- A running SOLIDWORKS session. The macro attaches to it and never starts a new instance.
- A saved production assembly. If no path is given, the active saved assembly is used. If several assemblies are open, activate the intended one or pass `sourcePath`.
- The PDM vault view and the latest kit assemblies in the sibling `12-KITS` folder of the product's `11-ROCK DRILL CADS` folder. A source outside PDM is matched by filename in an auto-login vault view.

## Outputs

Default output is a new folder beside the source, named `kitified-generated_<partno>_<yyyymmdd>_<run>`, where `<partno>` is the source file base name and `<run>` is the next free integer for that part number and date (starting at 1). It contains:

- The combined root `<source> kitified generated <run-id>.SLDASM`, with the generated children.
- A `Skycode_BoundingBox` feature saved in the combined root, plus `<combined root>_BoundingBox.SLDPRT` and `<combined root>_BoundingBox.txt`. Kitify creates the box in `Default` after configuration verification, using the canonical [bounding-box macro](./sldworks-bounding-box.md).
- Reports, including `run.txt`, `membership.csv`, `pdm-kit-comparison.csv`, `validation.csv`, `metadata.csv`, `configurations.csv`, and `kit-configurations.csv`.
- `declined-checkouts.txt`, listing PDM No responses queued by the helper or message-post failures.

Import the single combined root into Composer using the [canonical SOLIDWORKS import settings](<../../../docs/kitify sync.md#composer-solidworks-import-settings>), including all configurations. Do not import the generated kit children as separate masters.

## Side effects and restoration

- Temporarily changes SOLIDWORKS system options (save prompts, rebuild prompts, lightweight-component resolution) and restores the user's values on success or failure.
- Starts a hidden PowerShell helper for the run. It requests **No** on standard visible `SOLIDWORKS PDM` dialogs in the SOLIDWORKS process with exactly one No button, regardless of the question. Other dialogs are not answered; No may have consequences other than declining checkout. Logged responses are queued, not proof that the dialog accepted them. The updated matcher compiled and passed offline tests but has not been tested against a live dialog.
- Does not explicitly save or restructure the source assembly. Inspection can leave source documents dirty in memory; do not clear their dirty flags or discard pre-existing edits. The root Save As Copy does not copy referenced parts, which may remain shared. On failure partial output folders can remain.
- The older `kitify-generated_*`, `kitify-vault_*`, and `kitified-generated_*` folders are not deleted automatically.

## Gotchas

- If Composer reports `LitInteropSWCOM : Failed to load plug-in` and `OpenDoc failed`, check the [Composer SOLIDWORKS plug-in registration](<../../../docs/kitify sync.md#composer-solidworks-plug-in-registration>): after installing Composer 2026, the add-in still pointed to its 2026 DLL even after the SOLIDWORKS COM default was restored to 2025. Registering the 2025 plug-in resolved the observed import failure.
- Each kit configuration is named `<PartNo> <Description>` from the PDM kit assembly. Duplicate or reserved names fail the run.
- Create the root Save As Copy before resolving or inspecting source components; moving it later can trigger save prompts for modified references. This does not copy the referenced parts.
- Save Modified Documents prompts have also been observed during generation despite the existing prompt settings. Their cause remains unconfirmed; check `run.txt` around save, close, and reference-deletion calls rather than assuming the early-copy fix covers them.
- A source reference list previously showed `3115 2016 00 - Assemb.SLDASM` as missing. A Kitify `VERIFIED` result is not proof that every original reference exists; inspect references separately.
- Missing `PartNo` or `Description` on the source root, or on a kit after the NOTE fallback, stops the run. Identity is never invented from a filename.
- Suppressed components are excluded from kit membership but kept in `Default`.
- Boolean snapshots use `CBool(value <> False)`.
- CSV reports are comma-delimited, CRLF, UTF-8 with BOM.
- Partlink's BOM merge of kit CSVs is a separate step documented in [solidworks-pdm-bom-export](../../solidworks-pdm-bom-export/SKILL.md). It is not part of this macro.

## Latest live evidence

- On 2026-11-18, an isolated Access retry of local `7397 0500 28` completed in [run 4](<../../../partlink cad sandbox/kitified-generated_7397 0500 28_20261118_4/run.txt>). Its log reports 16 verified kit configurations and a saved bounding box; local source file sizes and modification times were unchanged. Both the initial and final source dirty flags were `True`.
- The PDM helper was deliberately disabled for that test, and the user reported one checkout dialog for `BG 006 019 67.SLDPRT`; the run was therefore not fully unattended. An earlier [run 1 checkout log](<../../../partlink cad sandbox/kitified-generated_7397 0500 28_20261118_1/declined-checkouts.txt>) records the previous helper responding to that same file. This does not verify the subsequently broadened matcher.
- The run logged PDM and child load warning code `2`; its meaning was not established. Do not equate this successful retry with a confirmed fix for the earlier save prompts or a guarantee of prompt-free future runs.
