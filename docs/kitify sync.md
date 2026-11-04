# Skycode kitify sync

## Current workflow

Working source: [Skycode - kitify.bas](../vba%20scripts/Skycode%20-%20kitify.bas).
This implements the revised Option D workflow, not donor editing or reinsertion
into the production drifter.

1. Attach to the running SOLIDWORKS session.
2. **First map the production tree**, recursively, including every exposed
   component occurrence and its full ancestry, referenced configuration,
   suppression/envelope status, NOTE and installed transform. Write the tree
   and a category/location inventory before learning or generating kits.
3. Read the separate open kitified reference as a hierarchy template.
4. Match source occurrences, learn category/scope ownership and plan independent
   kit and residual production assemblies.
5. Create new default-template assemblies inside-out, insert original source
   part files/configurations at installed positions, fix them and save them.
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

Open both saved assemblies in their intended configurations:

- Production: `[part number].sldasm`, for example `7393 0800 45.sldasm`.
- Reference: `[part number] kitified.sldasm`.

Run `SkycodeKitify`. Automatic selection uses a unique open production/reference
pair, not whichever child part happens to be active. Multiple matching pairs
are rejected instead of guessed.

For automation, call:

```vb
Dim result As String
result = KitifyDrifter(sourceFullPath, referenceFullPath, newOutputFolder, False)
```

All arguments are optional. Explicit paths must identify already-open assemblies.
An explicit output folder must **not already exist**; its parent must exist.
No permission, overwrite or output-location dialogs are used by the macro.
`showMessage=False` also disables the final informational message.

Default output:

```text
<source folder>\kitify-generated\<yyyymmdd_hhnnss>\
    <part number> kitified generated <run-id>.SLDASM
    KIT_<run-id>_<generic identifier>_<category>.SLDASM
    STRUCT_<run-id>_<identifier>.SLDASM
    ...reports...
```

Runs use separate folders and do not overwrite existing output. Keep the whole
run folder together: the root references its generated children. Run-specific
document names avoid SOLIDWORKS same-title conflicts with earlier open results.

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
- Promote each **topmost kit tree** out of its production ancestors to the new
  root. Preserve nested kit-to-kit relationships.
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
No PartNo, description or ERP metadata is assigned. `Skycode.*` properties on
new assemblies record kind, category, provenance and generated status.

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
Assembly-level geometry/modifications are not reconstructed. Shared source part
files retain their own dependencies; part-level in-context references are not
removed or certified. Production-document references are forbidden as
**assembly components** in the generated tree.

## Reports and acceptance criteria

Reports are UTF-16 CSV/text for Excel; commas, quotes and multiline values are
quoted. Transform values use the VBA host's locale and semicolon separators.

| File | Purpose |
|---|---|
| `source-tree.csv` | Full exposed production occurrence tree, before planning |
| `source-kit-locations.csv` | Category counts by full immediate-parent occurrence |
| `reference-tree.csv` | Reference snapshot |
| `membership.csv` | Source-instance destination, reference match and reasoning |
| `groups.csv` | Generated hierarchy, categories, provenance and descendant counts |
| `reference-kit-comparison.csv` | Reference versus planned source descendant quantities per learned kit |
| `comparison.csv` | Unmatched/ambiguous identities, aliases, NOTE differences and promotions |
| `generated-tree.csv` | Reopened generated occurrence tree |
| `validation.csv` | One source-to-generated occurrence match per retained part |
| `run.txt` | Stages, save/open warnings, dirty flags and explicit PASS/failure |

Success requires:

- Every active non-envelope source part occurs **exactly once**, including
  non-kit parts; no extra/reference-only parts.
- Original source file and referenced configuration are preserved.
- Each occurrence has its planned direct parent and fixed placement.
- Maximum absolute error in transform elements 0 through 12 is at most
  `1e-8` (translation elements are metres; rotation/scale are dimensionless).
- Every generated assembly component references the new run folder.
- All generated assembly parents match the plan and all assembly transforms
  are identity. Topmost kits are at root; nested kit parents match the reference.
- Save, reopen and rebuild succeed. Warnings are logged; failures raise an error
  identifying the stage and partial-output folder.

Source conservation PASS does **not** mean the production and hand-edited
reference contain identical hardware or geometry. Read the comparison report.
Partial outputs/unsaved generated documents can remain after an error; they are
not presented as a verified result.

## API basis

Uses the same typed COM connection, explicit Boolean comparisons, selection,
activation and fixed-component approach as the existing
[bounding-box macro](../vba%20scripts/Skycode%20-%20sldworks%20bounding%20box.bas)
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
- `ForceRebuild3(False)`, `IModelDocExtension.SaveAs3` (silent = 1).
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

Final live run, 4 November 2026, SOLIDWORKS revision `33.5.0`:

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

This is a full-snapshot generator. Persisted reusable rule files, incremental
regeneration, PDM integration and assigning service part numbers remain outside
this first implementation.
