# kitify

Macro: [Skycode - kitify.bas](<../../../vba scripts/Skycode - kitify.bas>)
Full specification: [kitify sync.md](<../../../docs/kitify sync.md>). Read it before changing this macro. This reference summarises it and does not replace it.

## Purpose

Builds one verified kitified assembly from the active production SOLIDWORKS assembly. The output is one root with `Default` plus one configuration per PDM kit assembly in the sibling `12-KITS` folder. Kit membership comes from the PDM kit assemblies, not from the source file.

## Host and entry points

- Runs in SOLIDWORKS VBA, or in Access with the SOLIDWORKS type library (`sldworks.tlb`) referenced. Access hosting is supported; the macro removes the OLE message filter during the run.
- Kitify itself does not add a bounding box. The combined entry and its reference are [kitify and bounding box](./kitifyandboundingbox.md).
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
- Reports, including `run.txt`, `membership.csv`, `pdm-kit-comparison.csv`, `validation.csv`, `metadata.csv`, `configurations.csv`, and `kit-configurations.csv`.
- `declined-checkouts.txt`, listing PDM No responses queued by the helper or message-post failures.

Import the single combined root into Composer using the [canonical SOLIDWORKS import settings](<../../../docs/kitify sync.md#composer-solidworks-import-settings>), including all configurations. Do not import the generated kit children as separate masters.

## Side effects and restoration

- Temporarily changes SOLIDWORKS system options (save prompts, rebuild prompts, lightweight-component resolution) and restores the user's values on success or failure.
- Starts a hidden PowerShell helper for the run. It requests **No** on standard visible `SOLIDWORKS PDM` dialogs (`#32770`) owned by the SOLIDWORKS process when exactly one button is No, regardless of the question. Other dialogs are not answered. No may have consequences other than declining checkout. A log line beginning `QUEUED NO` means the No command was posted, not that the dialog accepted it. This matcher is canonical: on 2026-11-19 it queued No on the live checkout dialog for `BG 006 019 67.SLDPRT` and the run continued without a manual answer.
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
- Kitify never adds a bounding box. Run `KitifyAndBoundingBox`, or add the box only after kitify returns and only while `Default` is active. `InsertEnvelope` resolves the new envelope in existing configurations, and deleting it from a kit configuration also deletes it from `Default`. Suppress the envelope and `Skycode_BoundingBox` in every non-Default configuration. `SetSuppression2(0)` success status is `2`. On the 2026-11-19 Composer import, Composer did not see the box in the kit configurations. The box stayed resolved in `Default`.
- Boolean snapshots use `CBool(value <> False)`.
- CSV reports are comma-delimited, CRLF, UTF-8 with BOM.
- Partlink's BOM merge of kit CSVs is a separate step documented in [solidworks-pdm-bom-export](../../solidworks-pdm-bom-export/SKILL.md). It is not part of this macro.

## Canonical live result (2026-11-19)

`KitifyAndBoundingBox` is the canonical way to produce a kitified root with a bounding box. The full record is in [kitify sync.md](<../../../docs/kitify sync.md#combined-kitify-and-bounding-box-live-run-2026-11-19>).

- Source `7397 0500 28.SLDASM` was copied, not saved. Size and last-write time stayed 13,789,862 bytes and 18/11/2026 8:40:47 PM. The in-memory dirty flag was already True and was left True.
- Output is [kitified-generated_7397 0500 28_20261119_1](<../../../partlink cad sandbox/kitified-generated_7397 0500 28_20261119_1/run.txt>). The date rolled, so this was run `_1`, not `_8`. Kitify verified `Default` plus 16 kit configurations.
- The PDM helper queued No on the live checkout dialog for `BG 006 019 67.SLDPRT` during kitify and again during the box reopen. See [kitifyandboundingbox-declined.txt](../../../logs/kitifyandboundingbox-declined.txt).
- The box was added after kitify, resolved only in `Default` (517 x 271 x 190 mm), and suppressed in every kit configuration. The user confirmed that Composer does not see the box in the kit configurations.
- PDM and child load warning code `2` was logged again. Its meaning is still not established. Do not treat this run as proof that every future prompt is impossible.

The 2026-11-18 [run 4](<../../../partlink cad sandbox/kitified-generated_7397 0500 28_20261118_4/run.txt>) is historical. Its PDM helper was disabled, so it does not supersede the 2026-11-19 decliner result.
