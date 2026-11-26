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
- **Save Modified Documents prompts (root cause analysis, 2026-11-26).** Kitify old 2 ran the 2026-11-19 canonical run with no prompts; the same code prompted repeatedly on 2026-11-26 (source `7377 2502 98+`). Do not "fix" this by copying references or by clicking Save all. What was established:
  - The prompts are raised by SOLIDWORKS for documents that are already dirty and are being closed. There are two dialogs: the list-style **Save Modified Documents** (raised when a kit is closed) and the task dialog **Save modified documents / "One or more documents being closed as a result of this operation have been modified"** (raised when `DeleteSelection2` clears the working assembly). The 2026-11-19 log shows neither.
  - The dirty documents are never generated ones: they are read-only vault kits and vault parts, and local Pack and Go subassemblies that the vault kits share by file name (`VS2550 DB-W+`, `7372 2900 80`, `7372 2592 45`, ...). `run.txt` lists them in `DIRTY DOCUMENT` lines before and after each kit close.
  - **Save all is wrong.** It saved the Pack and Go source files, and on the closing dialog it made SOLIDWORKS offer to save read-only vault parts, opening Save As dialogs into `C:\RDSA-PDM-Vault\...` (Cancel left the vault files unchanged). The answer must be **Don't Save**.
  - `CloseDoc` already closes dirty documents without saving, and no API clears a dirty flag (`SetSaveFlag` only sets it).
  - Fresh vault fetches were **not** the cause: all kits in the 2026-11-26 run were cached.
  - What differed from 2026-11-19: two kits (`7372 0986 01`, `7372 1650 02`) loaded with warning 34 (read-only + `swFileLoadWarning_NeedsRegen` = 32) and so open dirty; the 2026-11-19 kits all loaded with warning 2 (3 for one kit). Separately, the local Pack and Go subassemblies were dirty from the first kit onward and returned in the dialogs of later kits. A likely but unproven reason is that the vault kits reference subassemblies with the same file names as the source's, so SOLIDWORKS resolves them to the open local copies and rebuilds them; the log does not prove this. Neither difference is proven to be the whole cause, and no clean-session control run was made.
  - Copying every reference into the generated folder (`_skycopy_`/`_pdm_`, advanced Save As) was tried and abandoned: it produced 323 files instead of 37, changed canonical file and occurrence names, and still prompted.
  - **Verified outcome (run 10, 2026-11-26):** after SOLIDWORKS was restarted (a new process after the 22 GB session died) and the polluted source reference was repaired, the same macro passed end to end with **zero** save prompts and zero `QUEUED DONT SAVE` answers, even though `run.txt` logs 802 `DIRTY DOCUMENT` lines (read-only vault documents dirty in memory). Output matched the canonical shape: 55 files, 39 assemblies, no parts, no `_skycopy_`/`_pdm_` names, canonical `<base>_<runId>_A####` and `<PartNo> kitified <runId>` names. The only helper actions were the two usual PDM No answers. This supports the session-state explanation: prompts appeared in a long-lived SOLIDWORKS session that held stale in-memory source references from reference-copy experiments (plus a source subassembly saved with a bad reference). The helper's Don't Save answer is a safety net that was not needed in the clean run.
  - **Before any Kitify run after experiments or many runs:** restart SOLIDWORKS, confirm no source assembly references a generated folder (`GetDocumentDependencies2` on each source assembly, looking for `kitified-generated`), and check commit charge. SOLIDWORKS also warned "critically low on memory" at 44 of 45.7 GB commit with 9.4 GB RAM free (the `rm_commitcharge` warning counts virtual memory, not free RAM); a session that had run all day held 22 GB private memory and died during run 9. A fresh session peaked at about 13 GB for the full run.
  - **Mitigation in the macro:** the hidden PowerShell helper presses **Don't Save** (`BM_CLICK` on the button window; the command-link buttons have control ID 0, so a `WM_COMMAND` by ID triggers Save all instead) on any visible `Save Modified Documents` or `SOLIDWORKS` save dialog owned by the SOLIDWORKS process, for the whole run (flag file `skycode-kitify-discard-<run>.flag`). Generated documents are saved explicitly by the macro beforehand, so this never discards them. Answers are logged as `QUEUED DONT SAVE`.
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
