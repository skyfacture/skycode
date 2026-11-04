# Composer VBA Top Calls Cheat Sheet (Access Runtime)

## Scope
This cheat sheet is curated from:
- Real usage across this workspace's VBA modules.
- Extracted Composer API help topics under docs/swcomposerapi_html.

Use this file as a refactor guardrail. Prefer preserving call sequence and precondition checks.

## How Agents Should Use This
1. Identify the target behavior (views, metadata, traversal, transforms).
2. Find the matching section below.
3. Reuse the listed safe call sequence and guard checks.
4. Validate against Access runtime logs after any change.

## Top Usage Signals From Codebase
- Highest-frequency symbols in scripts: PartNo, Name, Selection, Meta, Actor, swDoc, View, Path, CLitSelection.
- Critical runtime objects/patterns: CLitSelection, document object usage (swDoc), property get/set flows, view create/rename flows, actor path parsing.

## 1) Document + Selection Access (CLitSelection)
### Primary Purpose
Get deterministic actor sets for traversal and operations.

### Typical Safe Sequence
1. Get document object once (swDoc).
2. Build or retrieve selection object (CLitSelection).
3. Enumerate selected/filtered actors.
4. Operate on actor identifiers without mutating selection mid-loop unless intentional.

### Guard Checks
- Ensure swDoc is not Nothing before any API call.
- Ensure selection count is validated before indexing.
- Avoid implicit reliance on UI current selection state; set/select explicitly when needed.

### Workspace Modules
- Skycode - selection only_subassemblies parts and nonnumeric partnumbers.bas
- Skycode - subassemblies parts and nonnumeric partnumbers.bas
- archive/Assy Selection Auto Views.bas

### API Topics
- docs/swcomposerapi_html/APIFiles/API-3DVIAComposer-Document.htm
- docs/swcomposerapi_html/APIFiles/API_AutoGenerate/group__group_composer_a_p_i_document_method.htm
- docs/swcomposerapi_html/APIFiles/API_AutoGenerate/group__group_composer_a_p_i_document_object.htm
- docs/swcomposerapi_html/APIFiles/API_AutoGenerate/group__group_composer_a_p_i_document_prop.htm

## 2) Actor Traversal by Path / Node Type
### Primary Purpose
Walk root, subassemblies, and parts in stable order.

### Typical Safe Sequence
1. Capture actor path string.
2. Normalize path (strip transient tokens, sanitize separators).
3. Infer hierarchy level (root/sub/part) from path segments and metadata.
4. Apply inclusion filter before expensive operations.

### Guard Checks
- Always sanitize path strings before comparisons.
- Keep hierarchy tests deterministic (same ordering each run).
- Log actor path + classification when skipping actors.

### Workspace Modules
- Skycode - subassemblies parts and nonnumeric partnumbers.bas
- Skycode - selection only_subassemblies parts and nonnumeric partnumbers.bas
- archive/Auto View Creation_Assy & Tree Order.bas

### API Topics
- docs/swcomposerapi_html/APIFiles/API_AutoGenerate/group__group_composer_a_p_i_document_object_prop.htm
- docs/swcomposerapi_html/APIFiles/API_AutoGenerate/group__group_composer_a_p_i_document_method.htm
- docs/swcomposerapi_html/APIFiles/FilterOperators.htm

## 3) Property Read/Write (Meta, PartNo, Description)
### Primary Purpose
Read source properties and write normalized metadata fields.

### Typical Safe Sequence
1. Read source property via document/object property call.
2. Sanitize/normalize value (strip prefix, trim, remove invalid tokens).
3. Write target property via set operation.
4. Re-read or log confirmation for critical properties.

### Guard Checks
- Null/empty checks before write.
- Keep sanitizers pure (no hidden side effects).
- Preserve exact metadata key names used by Composer profile.

### Workspace Modules
- Skycode - set meta partno with meta partno_cfg.bas
- Skycode - set meta description with Meta Description_cfg.bas
- Skycode - parts and nonnumeric partnumbers.bas

### API Topics
- docs/swcomposerapi_html/APIFiles/API_AutoGenerate/group__group_composer_a_p_i_document_object_prop.htm
- docs/swcomposerapi_html/APIFiles/API_AutoGenerate/group__group_composer_a_p_i_document_method.htm
- docs/swcomposerapi_html/APIFiles/HTMLProfiles.htm
- docs/swcomposerapi_html/APIFiles/HTMLInformationTemplateExample.htm

## 4) View Creation and Renaming
### Primary Purpose
Create views from actor groups and apply deterministic naming.

### Typical Safe Sequence
1. Compute target actor or actor set.
2. Create/select view operation.
3. Build final view name from sanitized business key.
4. Apply rename in sequence order.

### Guard Checks
- Ensure uniqueness of generated view names.
- Preserve sequencing assumptions (for example, starting offset like View 60).
- Log old and new view names for replay/debugging.

### Workspace Modules
- Skycode - rename views starting from View 60.bas
- Skycode - parts and nonnumeric partnumbers.bas
- archive/Skyfacture -- CreateViews 2_Sub Assemblies and Parts.bas

### API Topics
- docs/swcomposerapi_html/APIFiles/API_AutoGenerate/group__group_composer_a_p_i_document_method.htm
- docs/swcomposerapi_html/APIFiles/API_AutoGenerate/group__group_composer_a_p_i_document_prop.htm

## 5) Pivot / Frame Alignment and Local Coordinate Ops
### Primary Purpose
Align pivot actor data to frame/local coordinate expectations.

### Typical Safe Sequence
1. Fetch actor XML/property payload.
2. Extract required tags/attributes (Origin, Frame, Pivot).
3. Compare with tolerance.
4. Write corrected transform values.

### Guard Checks
- Use tolerance-based comparisons to avoid floating-point churn.
- Validate tag existence before numeric conversion.
- Keep format functions centralized and consistent.

### Workspace Modules
- Skycode - align pivot to actor location frame.bas
- Skycode - Set Each Pivot Actor Local Coord System.bas

### API Topics
- docs/swcomposerapi_html/APIFiles/API_AutoGenerate/group__group_composer_a_p_i_document_object_prop.htm
- docs/swcomposerapi_html/APIFiles/API_AutoGenerate/group__group_composer_a_p_i_document_method.htm

## 6) Batch/CLI and Export-Related Operations
### Primary Purpose
Use Composer command-line and output settings safely for automation.

### Typical Safe Sequence
1. Build file settings and input/output paths.
2. Select output format constants.
3. Execute and capture return/log info.

### Guard Checks
- Validate file paths before execution.
- Enforce expected output format constants.
- Log command parameters for reproducibility.

### API Topics
- docs/swcomposerapi_html/APIFiles/CommandLineParameters.htm
- docs/swcomposerapi_html/APIFiles/2DOutputFormatValues.htm
- docs/swcomposerapi_html/APIFiles/3DOutputFormatValue.htm
- docs/swcomposerapi_html/APIFiles/LogFileValues.htm
- docs/swcomposerapi_html/APIFiles/3DVIAConverterErrorCodes.htm

## Environment Troubleshooting

### Composer version upgrade breaks all scripts (2024 -> 2025)
Confirmed and resolved 2026-09-17. Composer 2025 is product version `7.12.4`.

**Symptom**
- Every script fails at `Set swApp = New DSComposer.Application`.
- `TestComposer` reports "Could not create Composer COM application using DSComposer.Application or Composer.Application".
- Tools > References shows no MISSING entries, and the project compiles, so the fault looks unrelated to references.

**Root cause**
The 2025 installer relocated the binary to `...\SOLIDWORKS Composer\bin\composer.exe` (2024 had it one level higher) and did not register the COM server. No Composer ProgID existed in `HKEY_CLASSES_ROOT`. The reference dialog appeared healthy because browsing directly to `composer.exe` reads the embedded type library from the file, which supplies compile-time types without requiring registration. Early binding therefore compiled cleanly and failed only at object creation.

**Diagnosis**
Run `DiagComposer` in tools/Skycode - composer connect diagnostic.bas, or check the registry directly:

```powershell
'DSComposer.Application','Composer.Application' | ForEach-Object {
  $c = (Get-ItemProperty ("Registry::HKEY_CLASSES_ROOT\" + $_ + "\CLSID") -EA SilentlyContinue).'(default)'
  if ($c) { $s = (Get-ItemProperty ("Registry::HKEY_CLASSES_ROOT\CLSID\" + $c + "\LocalServer32") -EA SilentlyContinue).'(default)'; "$_ = $c -> $s" }
  else { "$_ = NOT REGISTERED" }
}
```

**Fix**
1. Register the COM server from an elevated prompt:
   `"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS Composer\bin\composer.exe" /regserver`
   (Alternatively launch Composer once via Run as administrator.)
2. Copy the .accdb out of the OneDrive-synced folder before working on it; sync and `.laccdb` lock contention cause unrelated flakiness.
3. Decompile to clear P-code cached against the old registration:
   `"...\MSACCESS.EXE" "SkycodeHost SW v2025.accdb" /decompile`
4. Re-tick the Composer reference so it binds by GUID rather than file path, then Debug > Compile, then Compact and Repair.
5. Re-run `TestComposer`.

**Expected post-fix registry state**
- ProgID `DSComposer.Application` = `{3F41659D-BAB2-466C-A4D2-6E3CECFFB964}` -> `...\SOLIDWORKS Composer\bin\composer.exe`
- TypeLib `DSComposer` v1.0 = `{A71E1897-7597-441A-B1D8-725087CBF5FD}` -> same binary
- `Composer.Application` remains unregistered; 2025 publishes only the `DSComposer` ProgID, which is what the scripts use.

**Notes for future upgrades**
- The typelib GUID `{A71E1897-...}` is unchanged from 2024, so early-bound `Dim swApp As DSComposer.Application` declarations need no code changes across this upgrade.
- Check COM registration before suspecting API changes; a healthy-looking reference does not prove the server is registered.
- `New DSComposer.Application` / `CreateObject` spawns a fresh empty instance. If connection succeeds but `GetActiveDocument` returns Nothing (run-time error 91 downstream), attach to the running session with `GetObject(, "DSComposer.Application")` first. See `AttachComposer` in tools/Skycode - composer connect diagnostic.bas.

## Refactor Invariants (Do Not Break)
- Access is the execution host; do not shift runtime automation away from Access in refactor work.
- Preserve COM/API call ordering unless explicitly testing an alternative.
- Preserve actor-path normalization logic and hierarchy classification semantics.
- Preserve metadata key names and value sanitization behavior.
- Preserve logging semantics and summary counts.

## Suggested Agent Prompt Pattern
"Use COMPOSER_VBA_TOP_CALLS_CHEATSHEET.md and preserve the exact safe sequence for [section name]. Refactor only for readability and duplication reduction, no behavior changes."

## Related Workspace References
- CODE_SEMANTIC_INDEX.md
- COMPOSER_API_TOPIC_INDEX.md
- VBA_REFACTORING_BLUEPRINT.md
