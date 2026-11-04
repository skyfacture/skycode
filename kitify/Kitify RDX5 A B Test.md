```text
                                    RDX5 KIT SANDBOX A/B/C/D TEST
                                                 │
                                                 ▼
                                    Where should kit ownership live?
                                                 │
                 ┌───────────────────────────────┼───────────────────────────────┬────────────────────────────────────────────┐
                 ▼                               ▼                               ▼                                            ▼
             OPTION A                        OPTION B                        OPTION C                                   OPTION D
       Config-based kit views          Installed kit assemblies        Kits from kitified donor drifter       Kits synced from production drill
                 │                               │                               │                                            │
                 ▼                               ▼                               ▼                                            ▼
       Production Front Head           Production Front Head           Production drill unchanged             Production drill unchanged
        ├─ Seals                        └─ Front Head Kit              Create donor clone                     Single source of truth
        ├─ Fittings                         ├─ Seal Kit                Shared existing part files             Kit rules: scope + NOTE category
        └─ Fasteners                        ├─ Fastener Kit                      │                                            │
                                            └─ Hydr Kit                          ▼                                            ▼
       Existing configurations:                                       Complete Kits Assembly                  Sync macro (SolidWorks API / VBA)
       BEARING ONLY                    Production Impact Body          ├─ Front Head Kit                      Generated kit SLDASMs, no donor
       FASTENERS ONLY                   └─ Impact Body Kit             ├─ Gear Housing Kit                                    │
       HYDR FITTINGS ONLY                   ├─ Seal Kit                ├─ Impact Body Kit                                     ▼
                                            ├─ Fastener Kit            ├─ Bearing Kit                         Complete Kits Assembly
       Rename / identify using:             └─ Hydr Kit                   ├─ Ball bearing                      ├─ Front Head Kit
       Config Name = PN                                                   ├─ Bushing                           ├─ Gear Housing Kit
       Meta.PartNo = PN                Production Gear Housing            └─ Needle bearing                    ├─ Impact Body Kit
       Meta.Description                 └─ Gear Housing Kit            └─ Other Kits                           │  └─ Percussion Kit
                                            ├─ Seal Kit                         │                              ├─ Bearing Kit
                                            ├─ Fastener Kit                     ▼                              └─ Other Kits
                                            └─ Hydr Kit                 Insert into main drill (2 configs)                    │
                                                                                │                                             ▼
                 │                               │                      ┌───────┴────────┐                    In-situ transforms (Transform2)
                 ▼                               ▼                      ▼                ▼                    Change report -> review -> PDM
       Advantages                      Advantages                    Default            Kits                  Kits never edited by hand
       Minimal CAD change              Explicit ownership            Production ON      Production OFF                        │
       Fast implementation             Real BOM ownership            Kits OFF           Kits ON                               ▼
       Existing workflows              Natural kit hierarchy                    │                             Advantages
                                       Better Partlink integration              ▼                             Engineers edit drill only
                                                                        Common coordinate frame               No donor hardware
                 │                               │                      In-situ kit positions                 Cross-branch and nested kits
                 ▼                               ▼                      No move of installed parts            Repeatable, auditable sync
       Risks to test                   Risks to test                            │                                             │
       Governance clarity              Cross-branch kit membership              ▼                                             ▼
       ERP behaviour                   Composer hierarchy               Advantages                            Risks to test
       Metadata maintenance            Duplicate entities               Preserves production mates            Nested transform accuracy
                                       BOM / ERP behaviour              Supports cross-branch kits            Rule coverage / exceptions
                                                                        Discrete kit geometry                 Part configuration mapping
                 │                               │                              │                             PDM check-in / release
                 └───────────────┬───────────────┘                              ▼                                             │
                                 ▼                                      Risks to test                                         ▼
                    COMPLETE DRIFTER KIT OPTIONS                        Donor independence                    Reinsert into drill as config:
                    (original A/B proposal)                             Coordinate alignment                  deferred (Composer duplicates)
                                 │                                      Overlapping kit contents
                 ┌───────────────┴───────────────┐                      Membership maintenance
                 ▼                               ▼                      Circular references
       Envelope as configuration       Separate derived envelope                │
       in 7393 0800 45.SLDASM           COMPLETE KIT.SLDASM                     ▼
       Config: COMPLETE KIT            References governed kits         Kit / combination configs
                 │                               │                      in Complete Kits Assembly
                 └───────────────┬───────────────┘
                                 ▼
                    Subassembly kits own composition
                    Complete kit references those kits
                    Hardware edits occur at kit level

       ALL OPTIONS: SolidWorks structure / configs -> Composer views -> Partlink
```

## OPTION C - Donor-Based Service Kit Assembly

Added 23 October 2026. Options A and B remain in the diagram for
comparison. Option B is no longer the preferred approach where service-kit
members must remain inside different production subassemblies.

> Note: Option C produces the same kit structure as Option D, but the donor
> is maintained by hand and must be re-edited whenever the production drifter
> changes. Option D replaces the donor with a sync layer that generates the
> kit assemblies directly from the production drifter. See OPTION D below.

Option C separates the installed production hierarchy from the service-kit
hierarchy. The main drifter retains its production structure. An independent
donor assembly structure is reorganised into discrete kit assemblies, using
the same existing part documents and appropriate part configurations.

```text
Where should kit ownership live?
└─ OPTION C - Independent service-kit hierarchy
   ├─ Production assembly owns installed component relationships
   ├─ Individual kit assemblies own service-package membership
   └─ Complete Kits Assembly collects the individual kits
      └─ Inserted into the main drifter as an alternative configuration branch (non-kit geometry is deleted)
```

### SolidWorks structure and configurations

```text
7393 0800 45.SLDASM
├─ Original production subassemblies              [Default only]
├─ Original drill-level parts                     [Default only]
└─ Complete Kits Assembly.SLDASM                  [Kits only]
   ├─ Front Head service kits
   ├─ Gear Housing service kits
   ├─ Impact Body service kits
   ├─ 7372 1348 03 - RDX5 DRILL BEARING KIT.SLDASM
   │  ├─ 511 108 30 - Ball bearing x1
   │  ├─ 551 706 48 - Bushing x1
   │  └─ 888 097 49 - Needle bearing x1
   └─ Other service kits
```

The service-kit group labels above illustrate the separate hierarchy; they
do not require moving the corresponding installed components out of their
production subassemblies.

| Main drifter configuration | Original production components | Complete Kits Assembly |
|---|---|---|
| Default (no kits) | Unsuppressed | Suppressed |
| Kits (no production parts) | Suppressed | Unsuppressed |

"No production parts" means no original production-component instances are
active. Kit assemblies still contain their own instances of the shared parts.
Use configuration-specific suppression, not visibility alone.

### Donor workflow

1. Create an independent donor assembly structure from the main drifter.
   Retain references to the existing part documents rather than cloning their
   geometry. Copy child assembly documents wherever restructuring would
   otherwise modify a shared production assembly.
2. Establish a common coordinate-system reference for the main drifter,
   donor, complete-kits assembly and individual kit assemblies.
3. Build each service kit from the required donor component instances in
   their installed positions, including members drawn from different
   production branches. Preserve their position and orientation relative to
   the common frame. Coordinate systems provide an alignment reference, but
   placement must still be verified after restructuring.
4. Replace or remove donor mates that depend on components outside the new
   kit. Keep the production assembly's components and mates unchanged.
5. Collect the discrete kits in the Complete Kits Assembly. Insert that
   assembly into the main drifter, aligned to the common frame, and configure
   the mutually exclusive Default and Kits branches.

The final document-reference direction must be:

```text
Main Drifter -> Complete Kits Assembly -> Individual Kits -> Shared Parts
```

The Complete Kits Assembly must not contain a reference back to the main
drifter. Check inserted assemblies and in-context dependencies for circular
references before inserting it into the main drifter.

### Bearing-kit example

The installed bushing 551 706 48 remains inside BG 008 007 28, and the installed
needle bearing 888 097 49 remains inside BG 011 271 13. The separate bearing
kit contains new instances of those same part documents, together with
511 108 30. Kit creation changes service-package membership, not the installed
production hierarchy.

### Ownership and maintenance

- Part geometry is shared through the referenced part documents.
- Production subassemblies retain ownership of installed relationships.
- Individual service kits own kit composition and quantities.
- The Complete Kits Assembly references the individual kits; it does not
  independently maintain another list of loose kit hardware.
- The donor establishes the initial kit structure. Production changes do not
  automatically update kit membership, quantities or applicability; these
  must be reviewed and maintained.

Under Option C, a kit need not correspond to a physically installed
subassembly. The earlier requirement that only the complete drifter kit may
have a separate representation does not apply to cross-branch service kits.

### Overlapping kits

Different saleable kits can include the same replacement part. For example,
an Endurance Kit and a Bearing Kit may overlap. Showing both simultaneously
can therefore create overlapping geometry and repeated BOM quantities.

Use configurations in the Complete Kits Assembly to select an individual kit
or an intentional kit combination. The main drifter can retain just Default
and Kits, with its Kits configuration referencing the appropriate
Complete Kits Assembly configuration.

### Advantages

- Preserves production assembly structure and mates.
- Supports service kits whose parts occupy disparate installed locations.
- Provides discrete kit assemblies with physical geometry.
- Preserves in-situ positioning through a common coordinate frame.
- Keeps kit geometry available independently and through the main drifter's
  Kits configuration.

### Sandbox checks

- Confirm independent donor restructuring does not alter production assemblies.
- Verify shared part references, configurations and quantities.
- Verify kit geometry aligns with installed geometry when switching between
  Default and Kits.
- Confirm production and kit branches are mutually suppressed, with no
  unintended duplicate active geometry or BOM quantities.
- Confirm there are no circular assembly or in-context references.
- Test overlapping-kit selection and kit-specific BOM output.
- Test mates and exploded views after regrouping components.
- Verify how the chosen SolidWorks configurations determine Composer imports
  and view generation; Partlink then consumes those views.

## OPTION D - Sync Layer (Generated Service Kit Assemblies)

Added 23 October 2026. Option D keeps the Option C kit hierarchy but removes
the donor. Kit assemblies are generated from the production drifter by a
SolidWorks API macro, so engineering edits the drifter once and the kits
follow.

```text
Where should kit ownership live?
└─ OPTION D - Generated service-kit hierarchy
   ├─ Production drifter owns geometry, positions and installed structure
   ├─ Kit definitions own kit membership rules
   ├─ Sync macro generates the kit assemblies from the drifter
   └─ Complete Kits Assembly collects the generated kits
      └─ Reinsertion into the drifter as a config is deferred (see below)
```

### Principles

- The production drifter (7393 0800 45) is the single source of truth.
- Kit assemblies are generated outputs and are never edited by hand.
- No donor hardware: every kit part instance comes from the drifter's own
  part files, at its installed position.
- All kits share the drifter coordinate frame, so nested kits are inserted
  with an identity transform.

### Kit definitions

Most kits are defined by a rule: **scope assembly + NOTE category**. Explicit
lists cover kits that do not follow a category.

| Kit | Rule |
|---|---|
| 7372 1044 03 - RDX5 GH SEAL KIT | `SEAL KIT` parts under BG 009 287 40-1 |
| 7372 1045 03 - RDX5 IB SEAL KIT | `SEAL KIT` parts under BG 009 287 40-2 |
| 7372 1346 03 - RDX5 FH HYDR FITTING KIT | `HYDR KIT` parts under BG 009 205 35 |
| 7372 1348 03 - RDX5 DRILL BEARING KIT | `BEARING KIT` parts anywhere in 7393 0800 45 |
| 7372 2821 01 - RDX5 Percussion Kit | Explicit list (piston, guides, distributor, seals) |

The definitions file (Excel or CSV in PDM) holds one row per kit: kit PN,
description, scope assembly, category, explicit additions/exclusions,
quantity overrides and parent kit (for example Percussion Kit in Impact Body
Kit).

Rules are instance-based, not part-file-based. For example, O-ring 814 780 09
appears in both the Gear Housing and the percussion unit; scope rules keep
each instance in the correct kit.

### Generated structure

```text
Complete Kits Assembly.SLDASM                     [generated]
├─ Front Head service kits
├─ Gear Housing service kits
├─ Impact Body Kit
│  └─ 7372 2821 01 - RDX5 Percussion Kit.SLDASM
├─ 7372 1348 03 - RDX5 DRILL BEARING KIT.SLDASM
│  ├─ 511 108 30 - Ball bearing x1               (from BG 009 287 40-1)
│  ├─ 551 706 48 - Bushing x1                    (from BG 008 007 28)
│  └─ 888 097 49 - Needle bearing x1             (from BG 011 271 13)
└─ Other service kits
```

Nested kits are built inside-out: inner kits are generated and saved first,
then inserted into their parent kit.

### Sync workflow

1. The engineer edits the production drifter only and keeps `NOTE` correct.
2. The engineer runs Sync Kits (manual macro first; PDM workflow trigger
   later).
3. The macro walks the drifter, applies the kit rules and regenerates the
   affected kit assemblies at installed positions.
4. The macro produces a change report (added, removed, replaced, quantity
   changed, parent kits affected, unmatched parts).
5. The engineer reviews and approves the report; no manual kit edits.
6. The drifter and the changed kits are checked in and released together.
7. Composer regenerates the affected kit views; Partlink consumes them.

### SolidWorks API (verified against installed interop 33.5, SolidWorks 2025)

| Step | Call |
|---|---|
| Open drifter | `ISldWorks.OpenDoc6` |
| Traverse tree | `IAssemblyDoc.GetComponents`, `IComponent2.GetChildren`, `GetPathName`, `ReferencedConfiguration`, `IsSuppressed` |
| Read NOTE / properties | `ICustomPropertyManager.Get6` |
| Read installed position | `IComponent2.Transform2` |
| Create assembly | `ISldWorks.NewDocument` (default assembly template) |
| Insert parts at position | `IAssemblyDoc.AddComponents3` / `AddComponent5` |
| Build transforms | `ISldWorks.GetMathUtility`, `IMathUtility.CreateTransform` |
| Set kit metadata | `ICustomPropertyManager.Add3` / `Set2` |
| Kit combination configs | `IConfigurationManager.AddConfiguration2` |
| Fix in place | `IAssemblyDoc.FixComponent` |
| Save / close | `IModelDocExtension.SaveAs3`, `ISldWorks.CloseDoc` |

### Reinsertion into the main drifter

Reinserting the Complete Kits Assembly into 7393 0800 45 as a Kits
configuration is deferred. If reinserted, all kit parts are duplicated in
Composer. Evaluate after the Complete Kits Assembly exists. Circular
references are not a concern, because generated kits reference only part
files and other kit assemblies.

### Advantages

- Engineering edits the drifter once; kits follow automatically.
- No donor hardware to maintain.
- Supports cross-branch and nested kits.
- Repeatable, auditable sync with a change report.
- Physical kit trees for Composer view generation.

### Proof of concept and sandbox checks

1. Bearing kit only (3 parts across 3 nested subassemblies), manual macro.
2. Seal kits: scope rules, shared O-ring instances, nested GI / COMP kits.
3. Percussion Kit: explicit list, nested in Impact Body Kit.
4. PDM-triggered sync only after reports are reliably correct.

- Verify `Transform2` on deeply nested components gives drifter-level
  positions; otherwise compose transforms up the tree.
- Verify part configurations are mapped from the drifter instances.
- Verify the drifter is fully resolved (not lightweight) during sync.
- Verify PDM check-in of generated files.
- Verify Composer view generation from the generated kit assemblies.
- Decide ownership of the kit definitions file.
