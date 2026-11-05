# Skycode kitify sync

## Current workflow

Working source: [Skycode - kitify.bas]

1. Attach to the running SOLIDWORKS session.
2. **First map the production tree**, recursively, including every exposed
   component occurrence and its full ancestry, referenced configuration,
   suppression/envelope status, NOTE and installed transform. Write the tree
   and a category/location inventory before learning or generating kits.
3. Read saved reference membership/hierarchy rules; the development reference
   does not need to be open for normal runs.
4. Match source occurrences, learn category/scope ownership and plan independent
   kit and residual production assemblies.
5. Reuse **one working assembly document** to build assemblies inside-out.
   Insert original source parts/configurations at installed positions and fix
   them. Save child assemblies as copies, clear the working document, and reuse
   it for the next child and finally the root.
6. Close and reopen the saved generated assembly. Recursively verify occurrence
   conservation, configurations, transforms, fixed placement and kit parents.

The source and reference are never saved, restructured, assigned properties,
or used as generated assembly dependencies. Resolving lightweight components,
activating documents and reading uncached properties can affect their in-memory
state. The log records their initial/final dirty flags. Existing user edits are
not discarded.

## Running

Import the BAS into a standard VBA module in a SOLIDWORKS macro or Access.
Enable the installed **SOLIDWORKS type library** reference:
`C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\sldworks.tlb`.
No separate constants-library reference is required.

For one-time learning, open both saved assemblies in their intended configurations:

- Production: `[part number].sldasm`, for example `7393 0800 45.sldasm`.
- Reference: `[part number] kitified.sldasm`.

Run `LearnKitifyRules` once. It writes `<part number> kitify-rules.xml` beside the
source. Existing rules are not overwritten: move them aside explicitly before
relearning. The current drifter's rules were recovered from the previously
verified development run's reference snapshot.

For normal runs, close the development reference and leave only the source CAD
window open, then run `SkycodeKitify`. Keep the rules XML beside the source.
Automatic source selection accepts a unique production assembly with saved
rules (or a loaded development-reference partner); multiple candidates are
rejected instead of guessed. Missing/mismatched rules fail before generation.

Only the source and one working/final kitified assembly window are visible.
The same working document is reused for every child; no new SOLIDWORKS instance
is started. At most **two separately opened generated documents** exist during
construction: the working assembly and one hidden saved child being inserted.
That child is closed before loading another, below the requested maximum of
three. Children are saved with the Save As Copy option so the working document
keeps its identity. After each child is saved, the working document's components
are deleted to release references before building the next group.

Closing a document does not unload geometry still referenced by an assembly.
The source and final kitified assembly must retain their dependencies internally;
this is a bound on separately opened documents/windows, not all internally
loaded referenced files or a guarantee of constant memory usage.

For automation, call:

```vb
Dim result As String
result = KitifyDrifter(sourceFullPath, referenceFullPath, newOutputFolder, False)
```

All arguments are optional. Explicit CAD paths must identify already-loaded
assemblies. Leave `referenceFullPath` empty for normal saved-rule runs.
An explicit output folder must **not already exist**; its parent must exist.
No permission, overwrite or output-location dialogs are used by the macro.
`showMessage=False` also disables the final informational message.

Default output:

```text
<source folder>\kitify-generated_<yyyymmdd_hhnnss>\
    <part number> kitified generated <run-id>.SLDASM
    <part number> COMPLETE KIT.SLDASM
    <reference kit provenance tail>.SLDASM
    KIT_<run-id>_<generic identifier>_<category>.SLDASM
    <source non-kit assembly name>_<run-id>_<identifier>.SLDASM
    ...reports...
```

Runs use separate folders and do not overwrite existing output. Keep the whole
run folder together: the root references its generated children. Run-specific
root, structural and fallback-kit names avoid SOLIDWORKS same-title conflicts
with earlier open results. Reference-backed kits retain their provenance names;
close earlier results using those same kit titles before generating another run.
Only one output directory is created per invocation, directly beside the source
CAD. The older timestamp folders under `kitify-generated` are separate
development attempts, including partial runs; they are not automatically deleted.

Generated filenames and tree labels may differ from production. Non-kit backing
filenames start with the source assembly filename and include a unique run/group
identifier. Reference-backed kits use the segment after the final `/` in their
`groups.csv` provenance, or the whole value when there is no `/`, unchanged:
`7372 1045 03 RDX5 IB SEAL KIT-1.SLDASM`, for example. The provenance's final
occurrence suffix is retained deliberately; SOLIDWORKS adds its own instance
suffix in the tree (for example `7372 1045 03 RDX5 IB SEAL KIT-1-1`).
Reopen validation checks that tree names follow these filenames. Invalid
Windows filenames fail explicitly instead of being silently sanitized; duplicate
output filenames are rejected by the existing no-overwrite check.
Fallback kits have diagnostic provenance rather than a reference kit name and
keep the provisional `KIT_...` convention. Root and non-kit naming are unchanged.
No SOLIDWORKS-wide component-naming preference is changed.

All individual kit trees are children of one generated, uncoloured collection
at root level, named `<source root filename> COMPLETE KIT.SLDASM` (for example
`7393 0800 45 COMPLETE KIT.SLDASM`). Its `Skycode.Kind` is `COLLECTION`, not
`KIT`; it is a container, not another service kit. Existing nested kit
relationships are retained and parts are not duplicated. Residual production
assemblies stay outside the collection.

Each individual kit, including nested and fallback kits, receives an assembly
appearance from a deterministic non-grey palette: magenta, pink, green and cyan,
followed by darker and lighter shades of those four colours (12 distinct
colours total). No blue hues are used; cyan retains equal green/blue channels.
The palette cycles if there are more than 12 individual kits. These colours are applied
only to generated kit documents; original part/document appearances are not
edited. Appearance is cleared on the reused working document before every
group, so the collection, root and structural assemblies have no assembly
colour override. The saved/reopened appearance values are checked explicitly.

All generated CSV reports use comma delimiters, CRLF records and UTF-8 with a
BOM, not UTF-16/tab-delimited text. Text fields are double-quoted and embedded
quotes are doubled, preserving commas and Unicode in property values and names.
For Excel Text to Columns/import, select **Delimited**, **Comma** only, and the
double-quote text qualifier (UTF-8/65001 when an encoding choice is offered).
The diagnostic `run.txt` remains a separate Unicode text log.

The downstream identity contract is **properties**, not filenames/tree labels:
preserve `PartNo`, `Description` and, when present, `Global Part Number` on the
generated root and every rebuilt non-kit assembly. Read from the production
assembly and its referenced configuration, never the development reference.
Document-level properties remain document-level; configuration properties are
mapped to the generated snapshot's active configuration. Property spelling,
type and resolved value are copied; expressions are frozen to source-resolved
values so they do not re-evaluate against a generated filename.

This preserves the values consumed by Composer's `Meta.PartNo`/`Meta.PartNo_cfg`
and `Meta.Description`/`Meta.Description_cfg` normalization macros, including
their uppercase variants. Source part files/configurations are referenced
directly and their properties are not modified. New kits do not inherit a
production assembly's part number or description. The reused working document's
identity fields are cleared between groups to prevent metadata leakage.

Missing/blank source `PartNo` or `Description` is reported as `SOURCE_MISSING`
in `metadata.csv` and in the log; values are not invented from CAD filenames.
Such pre-existing gaps must be addressed before downstream publication.
Unmatched/ambiguous production owners for reference-learned physical
subassemblies fail explicitly rather than copying the reference's metadata.

## Membership and hierarchy rules

### NOTE reading

Read configuration-specific `NOTE` using `Get6` with uncached resolution; fall
back to document-level NOTE when blank. Match property names case-insensitively
through SOLIDWORKS. Normalize whitespace, case, hyphens, underscores and a
trailing KIT/KITS:

| Input examples | Category |
|---|---|
| SEAL, SEALS, SEAL KIT, SEALS KIT | SEAL |
| FASTENER, FASTENERS, FASTENER KIT | FASTENER |
| BEARING, BEARINGS, BEARING KIT | BEARING |
| HYDR, HYDR KIT, HYDR FITTING(S), HYDRAULIC FITTING(S) | HYDR FITTING |

Other nonblank NOTE values remain distinct categories; do not silently discard
them. Review unfamiliar categories in the reports. Membership is assigned to
**part occurrences**, not unique files: repeated hardware instances retain
their separate quantities, configurations and positions.

Suppressed components and envelopes, including all their accessible descendants,
are mapped but excluded from output. Hidden active non-envelope parts are included.
Suppressed assemblies may expose no children through `GetChildren`; this is an
active-configuration snapshot, not an inventory of every suppressed configuration.
Unresolved, unsaved/virtual or missing active components fail explicitly.

### Reference learning

- Detect kit assemblies by a separate KIT/KITS word in the filename, after
  removing `-kitified`. The word `kitified` alone is **not** a kit marker.
- Match part identity and referenced configuration, ignoring the reference's
  `-kitified` suffix. Embedded eight-digit or BG part-number aliases also permit
  differently named vendor CAD files to supply the reference grouping; alias
  matches are explicitly reported and do not certify geometric equivalence.
- Prefer matching configurations. An explicit part-number key may establish
  grouping despite a different reference configuration; log that difference and
  always retain the source configuration. Generic library files such as
  VORING/VWSHLOCK still require the same configuration, since it identifies the
  hardware variant.
- Prefer exact filenames and shared physical assembly ancestry. A second pass
  uses unambiguous matched peers to distinguish repeated hardware in different
  kit scopes. Ties across different kits remain explicit ambiguities rather
  than being resolved by first-match guessing.
- For matched members, reproduce the reference's internal kit hierarchy and
  physical subassemblies inside explicit kits. Examples include percussion
  and rotation kits, whose non-NOTE membership comes from the reference.
- Promote each **topmost kit tree** out of its production ancestors into the
  generated complete-kit collection at the new root. Preserve nested kit-to-kit
  relationships.
- Pool NOTE-tagged parts that the reference has not kitified using normalized
  category and the nearest unambiguous source-scope evidence. If ambiguous,
  create a generic category pool per top-level source occurrence. Unmapped
  bearings form one whole-drifter pool.
- A reference category kit cannot override a conflicting source NOTE category.
  An explicit mixed-content kit can contain different categories.
- Preserve remaining non-kit source parts in independently rebuilt production
  branches; prune branches with no retained active parts.
- Do not insert reference geometry to compensate for missing source hardware.
  Reference-only items and empty reference kits are reported, not manufactured.

Generic kit/structure identifiers are run-local, not service part numbers.
No new service part numbers are assigned. Existing production identity
properties are preserved on rebuilt non-kit assemblies; kits receive no
production identity. `Skycode.*` properties on new assemblies record kind,
category, provenance and generated status.

### Positions and assembly independence

All generated assemblies use the source drifter's coordinate frame. Source
occurrence `Transform2` arrays place the parts; generated child assemblies use
identity transforms. Do **not** multiply `Transform2` by ancestor transforms
again: traversal is through the root assembly's component context.

The first inserted component is unfixed before applying its transform. Reselect
after positioning, then fix it. SOLIDWORKS' VBA `Transform2` setter uses value
assignment, **without `Set`**; the installed COM implementation rejects
`Set component.Transform2 = ...`.

These are fixed geometry/configuration snapshots, not copies of production
mates, assembly features, exploded views, display states or all configurations.
Every generated component, including individual kits and the collection, is
fixed. Each generated assembly is checked for mates before saving and after
reopening; a nonempty mates group fails validation rather than being accepted.
Assembly-level geometry/modifications are not reconstructed. Shared source part
files retain their own dependencies; part-level in-context references are not
removed or certified. Production-document references are forbidden as
**assembly components** in the generated tree.

## Reports and acceptance criteria

CSV reports are UTF-8 with a BOM and comma delimiters; text values containing
commas, quotes or newlines are quoted. The separate `run.txt` log is UTF-16.
Transform values use the VBA host's locale and semicolon separators.

| File | Purpose |
|---|---|
| `source-tree.csv` | Full exposed production occurrence tree, before planning |
| `source-kit-locations.csv` | Category counts by full immediate-parent occurrence |
| `reference-tree.csv` | Learned reference membership snapshot; saved-rule runs use placeholder transforms, not reference geometry |
| `membership.csv` | Source-instance destination, reference match and reasoning |
| `groups.csv` | Generated hierarchy, categories, provenance and descendant counts |
| `reference-kit-comparison.csv` | Reference versus planned source descendant quantities per learned kit |
| `comparison.csv` | Unmatched/ambiguous identities, aliases, NOTE differences and promotions |
| `generated-tree.csv` | Reopened generated occurrence tree |
| `validation.csv` | One source-to-generated occurrence match per retained part |
| `metadata.csv` | Source versus reopened generated identity properties by scope/type; explicit source gaps |
| `appearance.csv` | Reopened individual-kit RGB colours; blank RGB for uncoloured collection/root/structures |
| `run.txt` | Stages, save/open warnings, dirty flags and explicit PASS/failure |

Success requires:

- Every active non-envelope source part occurs **exactly once**, including
  non-kit parts; no extra/reference-only parts.
- Original source file and referenced configuration are preserved.
- Each occurrence has its planned direct parent and fixed placement.
- All generated assemblies are mate-free, including the complete-kit collection.
- Every individual kit retains its assigned non-grey appearance after reopening;
  the collection, root and residual structures have no assembly colour override.
- Generation stays within the visible-window and separately opened generated
  document budgets above.
- Root/non-kit assembly identity properties retain their source scope, type and
  resolved values after save/reopen. Kits contain no leaked production identity.
- Maximum absolute error in transform elements 0 through 12 is at most
  `1e-8` (translation elements are metres; rotation/scale are dimensionless).
- Every generated assembly component references the new run folder.
- All generated assembly parents match the plan and all assembly transforms
  are identity. The complete-kit collection is at root; topmost individual kits
  are its children and nested kit parents match the reference.
- Save, reopen and rebuild succeed. Warnings are logged; failures raise an error
  identifying the stage and partial-output folder.

Source conservation PASS does **not** mean the production and hand-edited
reference contain identical hardware or geometry. Read the comparison report.
Partial outputs/unsaved generated documents can remain after an error; they are
not presented as a verified result.

## API basis

Uses the same typed COM connection, explicit Boolean comparisons, selection,
activation and fixed-component approach as the existing
[bounding-box macro](<../vba scripts/Skycode - sldworks bounding box.bas>)
and [local API notes](solidworksapi/best-fit-envelope.md).

Installed SOLIDWORKS 2025 interop signatures were inspected directly:

- `ISldWorks.GetFirstDocument`, `IModelDoc2.GetNext`, `ActivateDoc3`.
- `IAssemblyDoc.ResolveAllLightWeightComponents(False)`.
- `IConfiguration.GetRootComponent3(True)`, `IComponent2.GetChildren`.
- `GetPathName`, `GetModelDoc2`, `ReferencedConfiguration`, `IsSuppressed`,
  `IsEnvelope`, `IsFixed`, `Transform2`.
- `ICustomPropertyManager.Get6`, `Add3` (Text = 30; replace = 2).
- `GetUserPreferenceStringValue(9)` for the default assembly template.
- `NewDocument(template, 0, 0, 0)`.
- `AddComponent5(path, 0, "", True, sourceConfiguration, 0, 0, 0)`.
- `GetMathUtility.CreateTransform`, `Select4`, `UnfixComponent`, `FixComponent`.
- `ForceRebuild3(False)`, `IModelDocExtension.SaveAs3` (silent = 1;
  silent Save As Copy = 3), `IModelDoc2.Save3` (silent = 1, generated files only).
- `GetDocumentVisible`, `DocumentVisible`, `Visible`, `DeleteSelection2(0)`,
  `GetComponentCount(True)`; close inserted children before loading another.
- `ICustomPropertyManager.GetNames`, `GetType2`, `Get6`, `Delete2`, `Add3`
  for identity-property snapshot/copy/reopen verification.
- `CloseDoc`, `OpenDoc6` (assembly = 2; silent = 1).

## Validation status

Live validation uses the actual imported VBA module in a newly created isolated
Access database with the installed SOLIDWORKS type-library reference. No Access
security settings are changed. This avoids claiming that a C# approximation
establishes VBA runtime compatibility.

The initial source mapping found 304 exposed nodes and 23 NOTE/category
locations. The open reference is not an exact source clone: several kits remain
nested; some NOTE members remain loose; vendor filenames and some hardware
content differ. In particular, its three bearing-tagged parts include a bushing
that is not present as that part in the production snapshot. The macro must not
borrow that bushing from the reference.

Historical live run, 4 November 2026, SOLIDWORKS revision `33.5.0`:

- The actual VBA compiled; 12 normalization/identity assertions passed.
- Generated **36 SLDASM files**, including the root and **12 populated kits**.
- Closed and reopened the saved root, then verified **244/244 source part
  occurrences**, with no extras or missing occurrences.
- All 153 NOTE-tagged parts have category-correct kit ownership: 58 seals,
  47 fasteners, 46 hydraulic fittings and 2 bearings.
- Maximum transform error: `5.55111512312578e-16`, below `1e-8`.
- All generated components were fixed; all planned assembly parents and identity
  assembly frames passed. Both nested reference kit relationships were retained.
- An independent report check confirmed source occurrence/configuration/NOTE
  and active transforms were unchanged from the initial inspection.
- SHA-256 checks confirmed both source/reference root assembly files were
  unchanged on disk. Resolution/read operations did leave in-memory dirty flags.

The final comparison records 30 filename/part-number aliases, 3 configuration
differences (source configurations retained), 9 source-unmatched occurrences
(retained), and 4 reference-unmatched occurrences (not imported). Rotation kit
membership now matches all 6 reference members; percussion matches all 5.
There are also 28 deliberate NOTE promotions and one category correction:
the reference places a seal-tagged washer in its fastener kit. Kit quantity
differences are explicit in `reference-kit-comparison.csv`.

Verified output run folder: `kitify-generated\20261104_215023`.
Its root is
`7393 0800 45 kitified generated 20261104_215023_4AFB48B.SLDASM`.

This is a full-snapshot generator with persisted membership rules. Incremental
regeneration, PDM integration and assigning service part numbers remain outside
this implementation. Saved rules preserve development-reference membership,
not its geometry; unknown source occurrences still use the documented
NOTE/scope fallback and are reported for review.

### Current property-preservation validation, 5 November 2026

The actual VBA compiles. A four-part/four-child fixture passes save/reopen,
document-level `PartNo`, `Description` and `Global Part Number` preservation,
configuration-specific `Description` preservation, and kit metadata isolation.
It uses two visible windows at most and one hidden child document at a time,
without starting another SOLIDWORKS instance or changing naming preferences.

The full property-preserving retry is **verified** after restoring OneDrive
file access. Output: `kitify-generated_20261105_120116`, with root
`7393 0800 45 kitified generated 20261105_120116_2945A95.SLDASM`.

- Saved 36 assemblies: root, 23 rebuilt non-kit assemblies and 12 populated kits.
- Reopened normally with unchanged SOLIDWORKS naming preferences; verified
  244/244 source part occurrences and all fixed placements. Maximum transform
  error: `5.55111512312578e-16`, below `1e-8`.
- Independently compared 47 preserved identity properties to live production
  values and types after normal reopen; all passed. All 12 kits are free of
  inherited production identity.
- Reported 10 pre-existing source metadata gaps: 6 missing/blank `PartNo` and
  4 missing/blank `Description` entries in `metadata.csv`. These are not filled
  from filenames and still need review before publication.
- All 304 source nodes, configurations, NOTE values and active transforms match
  the original source snapshot; source root SHA-256 was unchanged by this run.
- Kept the original SOLIDWORKS process; only source and output windows remain.
  All 35 generated-child loads were closed before the next load, with a peak
  of two separately opened generated documents including the working assembly.

The earlier partial `kitify-generated_20261105_114632` run remains unaccepted:
Windows could not read its generated child or rules XML while the cloud file
provider was unavailable. That error alone did not establish CAD corruption.

### Provenance naming and comma-delimited reports, 5 November 2026

The updated VBA compiled in an isolated Access host and completed a full run
against the open `C:\skycode\partlink cad sandbox\7393 0800 45.SLDASM`.
Verified output: `kitify-generated_20261105_135039`, root
`7393 0800 45 kitified generated 20261105_135044_2F88FD4.SLDASM`.

- Saved and reopened 36 assemblies; all 244 source part occurrences, fixed
  placements and planned assembly parents passed.
- All nine reference-backed kit filenames and reopened tree names retain the
  final provenance segment, including its original occurrence suffix.
- All 12 kit descendant part identities and quantities still match the supplied
  `kitify-generated_20261105_120116\pdm kits.csv` baseline.
- Independently parsed all 10 CSV reports: UTF-8 BOM, comma delimiters and
  consistent expected field counts in every record.
- An actual VBA CSV fixture round-tripped commas, embedded quotes, Unicode,
  embedded CRLF and empty fields without changing values.
- All 47 preserved identity properties passed; the same 10 pre-existing source
  metadata gaps remain explicitly reported. No new SOLIDWORKS process was started.

### Complete-kit collection and individual-kit colours, 5 November 2026

Verified full output: `kitify-generated_20261105_142531`, root
`7393 0800 45 kitified generated 20261105_142531_3186A28.SLDASM`.

- Saved and reopened 37 assemblies, including the root-level uncoloured
  `7393 0800 45 COMPLETE KIT.SLDASM` collection.
- The collection contains 164 part occurrences across 12 individual kits;
  nested seal/fastener subkits remain inside their original kit parents.
- All 244 source parts remain conserved, with unchanged configurations and
  fixed placement. Every generated assembly is mate-free.
- All 12 individual kits retain distinct non-grey appearance colours after
  reopening. RGB values use SOLIDWORKS' persisted 8-bit channel precision.
- All 12 kit contents still match the supplied PDM CSV; all 11 generated CSV
  reports parse with consistent comma-delimited columns and UTF-8 BOMs.
- All 47 source identity properties still pass; the 10 pre-existing source
  metadata gaps remain reported. The source CAD disk hash is unchanged and the
  same SOLIDWORKS process was retained.

The earlier `kitify-generated_20261105_141621` test is unaccepted: its strict
appearance validation exposed 8-bit RGB rounding on mixed-channel colours.
The palette was corrected and the full run above passed without relaxing the
appearance comparison tolerance.

### Magenta/pink/green/cyan palette retest, 5 November 2026

Verified full output: `kitify-generated_20261105_150838`, root
`7393 0800 45 kitified generated 20261105_150838_33FE31B.SLDASM`.
Actual VBA compilation and complete save/reopen validation passed. The 12
individual kits retain three distinct shades each of magenta, pink, green and
cyan, with no blue hues; the complete-kit collection remains uncoloured.
All 244 source part occurrences remain fixed, every generated assembly is
mate-free, all 12 kit contents match the supplied PDM CSV, and all 11 CSV reports
have consistent comma-delimited fields and UTF-8 BOMs. The 47 preserved identity
properties pass; the same 10 pre-existing source metadata gaps remain reported.
The source disk hash is unchanged and the same SOLIDWORKS process was retained.
