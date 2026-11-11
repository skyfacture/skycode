# Skycode kitify sync

## Current workflow

Working source: [Skycode - kitify.bas](<../vba scripts/Skycode - kitify.bas>)

1. Attach to the running SOLIDWORKS session.
2. **Create the production-root copy before inspection**, then map the source
   tree recursively, including every exposed occurrence, ancestry, referenced
   configuration, suppression/envelope status, NOTE and installed transform.
3. Resolve the source product's own PDM `12-KITS` folder. A sandbox source is
   not a vault path: search auto-login vault views for the same filename
   directly in `11-ROCK DRILL CADS`, then use that product's sibling `12-KITS`
   folder. The canonical vault drifter is not opened, saved, or checked out.
   Read every SLDASM kit in that folder read-only, including its referenced
   files, and use each PDM assembly's part-occurrence multiset as the membership
   source. A kit already cached at its latest PDM version is read in place;
   it is not recopied while the source session has those files open. Match active source envelope NOTE values to PDM kits; do not use
   another drifter's model, reference tree or saved rules.
4. Reuse **one working assembly document** to build one generated branch for
   each PDM kit and the remaining non-kit production structures. A source
   occurrence may be reused in multiple kit branches when PDM lists it in
   multiple kits. Preserve its source configuration and transform in each
   membership branch.
5. Close and reopen the construction snapshot. Verify each generated kit
   branch's part multiset against its PDM assembly, including shared members,
   transforms, fixed placement, metadata and appearance.
6. In the production-root copy, preserve the original production tree, mates
   and features in `Default`. Add exactly one root configuration per SLDASM in
   the source product's PDM `12-KITS` folder. Each kit configuration activates
   only its corresponding generated kit branch; all other generated and
   original production branches are suppressed. A complete-kit configuration
   exists only when a complete-kit SLDASM exists in PDM. No synthetic
   complete-kit collection or configuration is created.
7. Reopen the combined output and verify the exact configuration inventory and
   each configuration's membership against PDM. Remove the intermediate
   construction snapshot only after successful verification.

The source assembly is never saved in place, restructured or assigned
properties. The source root is copied, not inserted as a component. `Default`
retains the original production child-document references. Every other
configuration corresponds to exactly one PDM kit assembly. No reference
assembly, learned rules file or source-specific membership file is read.
The production-root copy is created before resolving components or reading
uncached source properties, rather than late in the run after those operations
may have marked referenced documents modified in memory.
Resolving lightweight components, activating documents and reading uncached
properties can affect their in-memory state. The log records the source's
initial/final dirty flags. Existing user edits are not discarded.

## Running

Import the BAS into a standard VBA module in a SOLIDWORKS macro or Access.
Enable the installed **SOLIDWORKS type library** reference:
`C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\sldworks.tlb`.
No separate constants-library reference is required.

Open and activate the saved production assembly in SOLIDWORKS, then run
`SkycodeKitify`. The source may have any filename or reside in any folder.
A source outside PDM is matched to the same filename in an auto-login vault
view's `11-ROCK DRILL CADS` folder. Kits are read from the sibling `12-KITS`
folder. The production drifter itself is not opened or checked out.
No development reference assembly or adjacent rules file is needed. When no
source path is supplied programmatically, the active saved assembly is used; if
the active document is not an assembly, exactly one saved assembly must be open.
If several assemblies are open, activate the intended source or pass its full
path as `sourcePath`.

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
result = KitifyDrifter(sourceFullPath, newOutputFolder, False)
```

All arguments are optional. An explicit `sourcePath` must identify an
already-loaded assembly. An explicit output folder must **not already exist**;
its parent must exist.
No permission, overwrite or output-location dialogs are used by the macro.
`showMessage` defaults to `False`; pass `True` to show the final informational
message. `SkycodeKitify` never shows a message box: it prints the result or
failure to the Immediate window, and every run writes `run.txt`.

### Unattended runs (no dialogs)

For the duration of a run the macro temporarily changes these SOLIDWORKS system
options and restores the user's values on success or failure:

| Option | During the run |
|---|---|
| Don't prompt to save read-only referenced documents | On |
| Show errors every rebuild | Off |
| Warn before saving documents with update errors | Off |
| Warn before saving referenced documents | Off |
| Auto-dismiss messages while opening assemblies | On |
| Resolve lightweight components (when set to Prompt) | Always |
| Rebuild assembly on load (when set to Prompt) | Never |
| Rebuild on activation (when set to user decision) | Rebuild |

Documents are activated with an explicit rebuild instead of the rebuild prompt.
When hosted out of process (for example in Access), the host's OLE message
filter is removed for the run so
Access does not show its modal "waiting for another application to complete an
OLE action" dialog during long SOLIDWORKS calls; it is re-registered afterwards.
Inside SOLIDWORKS' own VBA host the filter is left alone.

The SOLIDWORKS PDM add-in asks "*file* is not checked out. Would you like to
check it out?" when a rebuild touches a read-only vault file. No SOLIDWORKS
preference or API call turns this prompt off, so for the duration of the run
the macro starts a hidden PowerShell helper (script in `%TEMP%`). The helper
answers **No** only to that PDM dialog in the SOLIDWORKS process. Nothing is
checked out. Each answered prompt is listed in `declined-checkouts.txt` and
copied into `run.txt` as `AUTO-DECLINED PDM CHECKOUT`.

The macro stops the helper and deletes its script when the run succeeds or
fails. The helper also exits by itself if SOLIDWORKS or the host process exits,
or if `run.txt` has not been written for 120 minutes.

Default output:

```text
<source folder>\kitified-generated_<yyyymmdd_hhnnss>\
    <source base name> kitified generated <run-id>.SLDASM
    <PDM kit identity> kitified <run-id>.SLDASM
    <source non-kit assembly name>_<run-id>_<identifier>.SLDASM
    ...reports...
```

The final root contains `Default` plus exactly one configuration for each PDM
kit SLDASM:

| Configuration | Original production branches | Generated branches |
|---|---|---|
| `Default` | Original suppression states, otherwise unchanged | All suppressed |
| `<PDM PartNo> <PDM Description>` | All suppressed | Only that PDM kit branch unsuppressed |

Originally suppressed production components remain suppressed in `Default`;
they are not forcibly unsuppressed. Hidden/envelope states, original occurrence
names, referenced configurations, placements and fixed/floating states are
retained. The final file is saved with the first PDM kit configuration active.
Import the **single combined root SLDASM** into Composer and select all
configurations to import `Default` and every PDM kit configuration. Do not
import individual generated kit SLDASMs as separate masters. Other source
configurations remain untouched in the original file but are not retained in
the output.

Runs use separate folders and do not overwrite existing output. Keep the whole
run folder together: the root references its generated children. Run-specific
root, structural and generated-kit names avoid SOLIDWORKS same-title conflicts
with earlier open results. Kit and structure names include the run identifier;
close earlier results using those titles before generating another run.
Only one output directory is created per invocation, directly beside the source
CAD. The older timestamp folders under `kitify-generated_*`,
`kitify-vault_*` and `kitified-generated_*` are separate development attempts,
including partial runs; they are not automatically deleted.

Generated backing filenames are run-specific and may differ from PDM filenames;
root configuration names use the corresponding PDM kit `PartNo` and
`Description`, separated by one space. These values come from the PDM kit
assembly's properties; a matched source-envelope NOTE only identifies the kit
and is used as a fallback if either PDM property is missing. Each kit
configuration does not carry configuration-specific `PartNo` or `Description`.
The generated kit child carries the effective kit identity as file-level
`PartNo` and `Description`, even when the PDM kit provides them at configuration
scope. The final root retains its production filename and file-level properties.
Its file-level `PartNo` and `Description` are expressions referring to the
configuration-specific `KitPartNo` and `KitDescription` helper properties.
Each kit helper contains the actual effective kit identity; `Default` helpers
contain the source's file-level production values. Composer imports the
resolved values into the ordinary `PartNo` and `Description` fields on each
configuration node. The helper fields may also appear separately.
No SOLIDWORKS-wide component-naming preference is changed.

Each individual PDM kit branch receives an assembly appearance from a
deterministic non-grey palette: magenta, pink, green and cyan,
followed by darker and lighter shades of those four colours (12 distinct
colours total). No blue hues are used; cyan retains equal green/blue channels.
The palette cycles if there are more than 12 individual kits. These colours are applied
only to generated kit documents; original part/document appearances are not
edited. Appearance is cleared on the reused working document before every
group, so the construction snapshot and structural assemblies have no assembly
colour override. The final cloned root retains its production appearance. The
saved/reopened generated-child appearance values are checked explicitly before
combination.

All generated CSV reports use comma delimiters, CRLF records and UTF-8 with a
BOM, not UTF-16/tab-delimited text. Text fields are double-quoted and embedded
quotes are doubled, preserving commas and Unicode in property values and names.
For Excel Text to Columns/import, select **Delimited**, **Comma** only, and the
double-quote text qualifier (UTF-8/65001 when an encoding choice is offered).
The diagnostic `run.txt` remains a separate Unicode text log.

The downstream identity contract is **properties**, not filenames/tree labels:
preserve all source custom properties, including `PartNo`, `Description` and,
when present, `Global Part Number`, on the generated root and every rebuilt
non-kit assembly. The four macro-owned `Skycode.Kind`, `Skycode.Category`,
`Skycode.Provenance` and `Skycode.Generated` fields are reserved, not copied.
Read from the production assembly and its referenced configuration.
Document-level properties remain document-level; configuration properties are
mapped to each generated snapshot's active configuration. Kit `PartNo` and
`Description` are the exception: their effective values are written at file
scope on the generated kit child. Root kit configurations must not contain
configuration-specific `PartNo` or `Description`. The cloned `Default`
configuration retains its production
configuration properties. Property spelling, type and resolved value are copied;
source expressions are frozen to source-resolved values so they do not re-evaluate
against a generated filename. The combined root's two identity expressions are
the intentional exception and resolve against the active configuration's helpers.
The source root's file-level `PartNo` and `Description` must be nonempty text
properties. The run stops if either kit identity is missing after the NOTE
fallback, or if source root metadata already uses either reserved helper name;
it never invents Composer identity from a filename.

This preserves the values consumed by Composer's `Meta.PartNo`/`Meta.PartNo_cfg`
and `Meta.Description`/`Meta.Description_cfg` normalization macros, including
their uppercase variants. Source part files/configurations are referenced
directly and their properties are not modified. The reused working document's
source custom properties are cleared between groups to prevent metadata leakage.

For a kit with a matched source envelope, its NOTE identifies the matching PDM
kit but does not override the PDM kit's `PartNo` or `Description`. A missing
PDM property may be filled from the NOTE; otherwise the PDM kit's own values
are used. Each combined-root kit configuration is named
`<PartNo> <Description>`. Its generated kit child stores the effective identity
at file scope; the root kit configuration does not store these properties at
configuration scope. No synthetic complete-kit properties are added.

Composer import must include custom properties and use `PartNo` for actor names
where configured. The root
view/BOX name still comes from root `Actor.Name`; correct root naming therefore
depends on the Composer import naming settings. Existing imports must be
re-imported/updated to receive the new metadata. BOX views additionally require
the bounding-box macro's proxy and its geometry properties; kit metadata alone
does not create that proxy.

Missing/blank `PartNo` or `Description` on other source assemblies is reported
as `SOURCE_MISSING` in `metadata.csv` and in the log; values are not invented
from CAD filenames. Such pre-existing gaps must be addressed before downstream
publication. Missing required root or kit identity stops the run, with details
in the log.

## PDM kit and configuration rules

### PDM source of truth

The source assembly must be inside a product's `11-ROCK DRILL CADS` folder.
Kitify resolves only the direct sibling `12-KITS` folder in that same PDM
product folder. Every SLDASM in that folder is represented by exactly one
configuration named `<PartNo> <Description>` from the PDM kit assembly's
properties. Duplicate or reserved configuration names fail rather than being
renamed silently.

The PDM assembly's active leaf part-occurrence multiset defines kit contents.
Repeated parts preserve quantity. A part can appear in more than one kit; it is
rebuilt in each corresponding generated branch, and each root configuration
activates only its matching branch. A complete-kit configuration is generated
only when PDM contains that assembly. The macro never infers a union kit,
collection, subset hierarchy, or additional configuration.

Active envelope NOTE values are used to match the source envelope to a PDM kit.
A missing or ambiguous envelope match fails explicitly. The PDM kit's own
`PartNo` and `Description` are authoritative; its NOTE is only a fallback for
a missing property. PDM kits without a matched envelope still get their
canonical configurations and use their own PDM properties.

Suppressed components and envelopes, including all their accessible descendants,
are mapped but excluded from the generated kit tree. They remain intact in the
cloned production tree. Hidden active non-envelope parts are included in kits.
Suppressed assemblies may expose no children through `GetChildren`; this is an
active-configuration snapshot, not an inventory of every suppressed configuration.
Unresolved, unsaved/virtual or missing active components fail explicitly.

### Production preservation and unmatched parts

- Suppressed components and envelopes, including accessible descendants, are
  mapped but excluded from generated kit membership; they remain intact in
  `Default`. Hidden active non-envelope parts are included when the PDM kit
  assembly requires them.
- Active source parts that occur in no PDM kit are retained in independently
  rebuilt production structures, following their source ancestry.
- If a PDM kit requires more occurrences of a part than the active source tree
  provides, the missing occurrences are placed from that PDM kit assembly and
  logged as `KIT_SUPPLIED`. A kit file that is not already loaded is opened
  read-only and hidden, then inserted. It is not saved or checked out. Generation
  still fails if the PDM kit snapshot itself cannot supply the required identity.
  It does not silently drop parts.

Generated kit and structure filenames are run-local. Existing production
custom properties are preserved on rebuilt non-kit assemblies; matched kits
receive provenance-derived identity, and unmatched kits preserve PDM identity.
`Skycode.*` properties on new assemblies record kind, category, provenance and
generated status.

### Positions and assembly independence

All generated assemblies use the source drifter's coordinate frame. Source
occurrence `Transform2` arrays place the parts; generated child assemblies use
identity transforms. Do **not** multiply `Transform2` by ancestor transforms
again: traversal is through the root assembly's component context.

The first inserted component is unfixed before applying its transform. Reselect
after positioning, then fix it. SOLIDWORKS' VBA `Transform2` setter uses value
assignment, **without `Set`**; the installed COM implementation rejects
`Set component.Transform2 = ...`.

The generated kit branches are fixed geometry/configuration snapshots, not
copies of production mates, assembly features, exploded views, display states
or all configurations. Every generated kit-tree component is fixed. Each
generated assembly is checked for mates before saving and after
reopening; a nonempty mates group fails validation rather than being accepted.
The final cloned root is exempt: it deliberately retains the original production
mates/features. Original production occurrences keep their fixed/floating state.
Assembly-level geometry/modifications are not reconstructed. Shared source part
files retain their own dependencies; part-level in-context references are not
removed or certified. Production-document references are forbidden as
**assembly components** in the active generated kit tree, but are required and
preserved in `Default`.

## Reports and acceptance criteria

CSV reports are UTF-8 with a BOM and comma delimiters; text values containing
commas, quotes or newlines are quoted. The separate `run.txt` log is UTF-16.
Transform values use the VBA host's locale and semicolon separators.

| File | Purpose |
|---|---|
| `source-tree.csv` | Full exposed production occurrence tree, before planning |
| `source-kit-locations.csv` | Category counts by full immediate-parent occurrence |
| `membership.csv` | Source occurrence destination and all PDM kit configurations containing it |
| `groups.csv` | Generated PDM kit/structure branches, configuration names and counts |
| `pdm-kit-comparison.csv` | PDM part-identity quantities compared with generated source assignments |
| `generated-tree.csv` | Reopened generated occurrence tree |
| `snapshot-generated-tree.csv`, `snapshot-validation.csv`, `snapshot-metadata.csv` | Construction-snapshot verification before combining; separate names preserve the no-overwrite report policy |
| `production-tree.csv` | Reopened `Default` tree, including suppressed generated branches |
| `production-metadata.csv` | Reopened `Default` root custom properties compared with source-resolved values |
| `configurations.csv` | Exact root configuration inventory, branch suppressions and per-configuration PDM quantities |
| `kit-configurations.csv` | PDM PartNo/Description, configuration name, resolved ordinary identity and verification that configuration-specific identity properties are absent |
| `declined-checkouts.txt` | PDM "check it out?" prompts answered No automatically (only present when one occurred) |
| `validation.csv` | Source-to-generated occurrence and transform checks, including reused memberships |
| `metadata.csv` | Expected versus reopened properties by scope/type; source and PDM-kit metadata gaps |
| `appearance.csv` | Reopened individual-kit RGB colours; blank RGB for root and structures |
| `run.txt` | Stages, save/open warnings, dirty flags and explicit PASS/failure |

Success requires:

- The final root contains `Default` and exactly one configuration per SLDASM in
  the source product's PDM `12-KITS` folder; no extra kit configurations exist.
- Each PDM kit configuration contains exactly the part identities and
  quantities in its PDM assembly. Shared members can recur across
  configurations, never as an inferred or synthesized membership.
- A complete-kit configuration exists only when that complete-kit SLDASM is
  present in PDM.
- Original source file and referenced configuration are preserved.
- Each occurrence has its planned direct parent and fixed placement.
- All generated child assemblies are mate-free; the cloned root retains the
  original production mates/features.
- Every individual kit retains its assigned non-grey appearance after reopening;
  residual structures have no assembly colour override.
- Generation stays within the visible-window and separately opened generated
  document budgets above.
- Root/non-kit assembly custom properties retain their source scope, type and
  resolved values after save/reopen in `Default`. The root's file-level identity
  expressions resolve to each kit's actual identity in that kit's configuration.
- Root kit configurations contain no configuration-specific `PartNo` or
  `Description`; kit child identities are file-level properties.
- `Default` retains every exposed production occurrence, filename, referenced
  configuration, suppression/envelope status, fixed/floating state and placement.
  In a PDM kit configuration, only its selected kit branch is active; production
  and other generated branches are suppressed.
- Maximum absolute error in transform elements 0 through 12 is at most
  `1e-8` (translation elements are metres; rotation/scale are dimensionless).
- Every active generated assembly component in a PDM kit configuration references the new
  run folder; the production branches retain their original references.
- Generated assembly parents match the plan and all generated assembly
  transforms are identity. Each PDM kit branch is a root child; membership and
  kit configuration names come from the source-relative PDM folder.
- Save, reopen and rebuild succeed. Warnings are logged; failures raise an error
  identifying the stage and partial-output folder.

Source conservation PASS verifies that the generated tree matches the source
plan; it does not certify that NOTE values are semantically correct. Review
`membership.csv` and `pdm-kit-comparison.csv`. Partial outputs/unsaved generated
documents can remain after an error; they are not presented as a verified result.

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
  for custom-property snapshot/copy/reopen verification.
- `AddConfiguration3`, `DeleteConfiguration2`, `IConfiguration.Name`,
  `ShowConfiguration2` and `GetComponents(True)`.
- `IComponent2.SetSuppression2`: suppressed = 0, fully resolved = 2;
  success (`swSuppressionChangeOk`) = 2. Changes apply to the active configuration.
  State acceptance uses the reopened configuration checks.
- PDM vault resolution uses `GetVaultNameFromPath` when the source is already
  inside a vault view. A non-vault path returns `0x8004022B` (`-2147220949`,
  "The key you specified was not found in the collection") and is not itself a
  missing-collection bug. Sandbox sources then search each auto-login local
  vault view for the same filename directly under `11-ROCK DRILL CADS`, with
  sibling `12-KITS` confirmed through `GetFolderFromPath`. Exactly one match is
  required. `IEdmFolder5.GetFirstFilePosition` / `GetNextFile` enumerate only
  that product's direct `12-KITS` folder. `IEdmFile5.GetFileCopy` receives
  the latest file version, folder ID, `EdmGet_MakeReadOnly Or EdmGet_Refs`
  (flags 1 and 4), and no rename; reference files are fetched at their PDM
  assembly references. The assembly's local version is checked before use.
- Normalize native Boolean flags with `CBool(flag <> False)` before comparing
  them with VBA Boolean values. SOLIDWORKS can return true as `1`, not VBA's
  canonical `True` (`-1`); `CBool(flag)` alone does not normalize a typed Boolean.
- Configuration activation is skipped when the requested configuration is
  already active, since `ShowConfiguration2` can return False for that no-op.
  After an actual switch, rebuild the activated configuration before checking
  its geometry and component tree.
- `CloseDoc`, `OpenDoc6` (assembly = 2; silent = 1).

## Validation status

The user confirmed that an expression fixture with actual kit values imported
into Composer with correct ordinary `PartNo` and `Description`. A newly
generated output should also be imported and its views and exploded BOMs
checked before publication.

Live validation uses the actual imported VBA module in a newly created isolated
Access database with the installed SOLIDWORKS type-library reference. No Access
security settings are changed. This avoids claiming that a C# approximation
establishes VBA runtime compatibility.

Run `TestKitifyMetadata` in the imported module to verify the NOTE parsing
fallback. Full generation must verify `Default`, every saved PDM kit
configuration after reopening, and that each configuration's name and
configuration-specific `PartNo`/`Description` are absent, the root expressions
resolve to the expected values, and the generated kit child retains its
effective file-level identity. The checks
are captured in `configurations.csv`, `kit-configurations.csv`,
`pdm-kit-comparison.csv`, `production-tree.csv`, `validation.csv`, and
`metadata.csv`.

### Expression identity, live run (2026-11-11)

The actual VBA module compiled and passed its metadata and identity-plan guard
tests in an isolated Access host. A full generation from the open
`7397 0320 22.SLDASM` completed in `kitified-generated_20261111_155750`.
After reopening, all 16 kit identities and the production `Default` identity
resolved correctly; kit configuration `PartNo`/`Description` were absent.
The production tree and PDM membership checks also passed. The source was not
saved, but remained dirty in memory. Composer import of this generated run and
its view/BOM workflow remain to be verified.

### Kit configuration identity scope, live run (2026-11-11)

The actual VBA module passed `TestKitifyMetadata` in an isolated Access host.
A full retry against the open `7397 0320 22.SLDASM` completed and reopened its
output in `kitified-generated_20261111_110210`. All 16 rows in
`kit-configurations.csv` passed with configuration-specific identity properties
absent. The generated children for `7372 3010 01` and `7372 3011 01` retained
their kit `PartNo` and `Description` at file scope, as reported in
`snapshot-metadata.csv`. `Default` and PDM membership validation also passed.

The standard `GetObject(, "SldWorks.Application")` connection failed in that
session. The isolated test host supplied the existing SOLIDWORKS process via
its process-specific running-object registration, substituting only the
connection statement in its imported in-memory module; the repository module's
connection code was not changed. The source was not saved, although it was
dirty in memory after the run. Composer identity and BOM behaviour have not
been verified by this run.

### PDM-named kit configurations, unattended run (2026-11-07)

The actual VBA module was imported into an isolated Access host and run against
the local `7397 0320 22.SLDASM` drifter. It ran to completion with no manual
input. Output: `kitified-generated_20261107_000811`, verified root
`7397 0320 22 kitified generated 20261107_000811_78087.SLDASM`.

- All 16 PDM kits became root configurations named `<PartNo> <Description>`
  (for example `550 361 58 RD106 FH SEAL KIT`). Every row in
  `kit-configurations.csv` passed.
- 485 occurrences were conserved, and the hierarchy, metadata, appearance, mate
  and `Default` production-tree checks all passed.
- During configuration validation, SOLIDWORKS PDM asked to check out
  `864 095 09.SLDPRT`. The helper answered No automatically and the answer was
  logged in `declined-checkouts.txt` and `run.txt`.
- The user's SOLIDWORKS preferences were restored afterwards. The helper process
  exited and its temporary script was deleted.

A run made after many earlier runs in the same SOLIDWORKS session failed to
reopen its saved output with `swFileRequiresRepairError` (2097152). In that
state SOLIDWORKS also refused files that had opened correctly before. After
SOLIDWORKS was restarted, the run above passed. If a long session reports this
error, restart SOLIDWORKS before running again.

### Combined-configuration refinement (2026-11-05)

The actual VBA module compiled and its metadata regression tests passed.
A complete construction run produced and verified the kit snapshot; targeted
combination and cold-reopen checks then verified the final saved root:

- Both `Default` and `kitified` passed all 11 top-level suppression checks.
- All 304 exposed production occurrences retained their source paths,
  configurations, suppression/envelope states, fixed/floating states and
  placements in `Default`.
- All 244 retained parts were conserved in `kitified`; maximum matrix error was
  `5.69544411632705e-14`, below the required `1e-8` tolerance.
- All 417 generated custom-property checks passed, including
  `PartNo = 7372 1043 01` for the example seal kit. The `Default` root also passed
  all 22 source-property checks.
- Five unnumbered kits and ten pre-existing source identity gaps remain
  explicitly reported; no numbers or missing production values were invented.
- The intermediate construction assembly was removed after verification.
- The source disk SHA-256 remained unchanged. Resolution/read operations left
  the source dirty in memory; the macro did not save that source state.

The live Boolean diagnostic confirmed native true was `1`, that
`CBool(IsSuppressed) = True` was False, and that
`CBool(IsSuppressed <> False) = True` was True. Configuration verification uses
the existing normalized-Boolean pattern throughout.

These checks verify SOLIDWORKS output and its metadata, not a new Composer
import or creation of BOX views. Re-import the intended configuration with the
custom-property naming settings and supply/regenerate the bounding-box proxy.

### Earlier snapshot-only validation

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

That earlier version was a full-snapshot generator with persisted membership
rules. The current version no longer reads those rules or a development
reference; it classifies from source NOTE/category and source scope only.
Incremental regeneration, PDM integration and assigning service part numbers
remain outside this implementation.

### Historical property-preservation validation, 5 November 2026

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

### Historical provenance naming and comma-delimited reports, 5 November 2026

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

### Historical complete-kit collection and individual-kit colours, 5 November 2026

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

### Historical magenta/pink/green/cyan palette retest, 5 November 2026

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
