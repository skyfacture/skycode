---
name: partlink-solidworks-macros
description: 'Develop, edit, run, choose, or troubleshoot the Skycode SOLIDWORKS and SOLIDWORKS Composer VBA macros used in Partlink maintenance and operation (kitify, bounding boxes, views, metadata, pivots, explode, root identification).'
---

# Partlink SOLIDWORKS and Composer macros

The macros in [vba scripts/](<../../vba scripts>) support Partlink maintenance. Each `.bas` file is imported into a standard VBA module. Each macro has a reference file in [references/](./references) with its entry point, inputs, outputs, side effects, and gotchas. Read the reference before running a macro, and read the `.bas` file before changing one.

## Hosts

- **SOLIDWORKS VBA** (attaches to the running SOLIDWORKS session): [kitify](./references/kitify.md), [kitify and bounding box](./references/kitifyandboundingbox.md), [sldworks bounding box](./references/sldworks-bounding-box.md).
- **Composer VBA** (uses `DSComposer`, run from an Access host that references `C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS Composer\bin\composer.exe`): every other macro. Do not late-bind the Composer COM object from PowerShell; test Composer VBA through the Access host.

## Macro index

| Macro | Host | Entry point | Reference |
|---|---|---|---|
| kitify | SOLIDWORKS | `SkycodeKitify`, `KitifyDrifter` | [kitify](./references/kitify.md) |
| kitify and bounding box | SOLIDWORKS | `KitifyAndBoundingBox` | [kitifyandboundingbox](./references/kitifyandboundingbox.md) |
| sldworks bounding box | SOLIDWORKS | `SldworksBoundingBox` | [sldworks-bounding-box](./references/sldworks-bounding-box.md) |
| bounding box dims | Composer | `BoundingBoxDimensions` | [bounding-box-dims](./references/bounding-box-dims.md) |
| views [assemblies + all nonnumeric parts + bounding box dims] | Composer | `CreateViewsAllWithBoundingBoxDims` | [views-assemblies-all-nonnumeric-parts-bbox-dims](./references/views-assemblies-all-nonnumeric-parts-bbox-dims.md) |
| views [assemblies + all nonnumeric parts] | Composer | `CreateAssemblyViews` | [views-assemblies-all-nonnumeric-parts](./references/views-assemblies-all-nonnumeric-parts.md) |
| views [assemblies + top-level nonnumeric parts] | Composer | `CreateAssemblyViews` | [views-assemblies-top-level-nonnumeric-parts](./references/views-assemblies-top-level-nonnumeric-parts.md) |
| views [parts-only nonnumeric] | Composer | `CreateAssemblyViews` | [views-parts-only-nonnumeric](./references/views-parts-only-nonnumeric.md) |
| views [create balance parts not in db] | Composer | `CreateViewsForPartsNotInDb` | [views-create-balance-parts-not-in-db](./references/views-create-balance-parts-not-in-db.md) |
| views [delete parts already in db] | Composer | `DeleteExtraneousViews` | [views-delete-parts-already-in-db](./references/views-delete-parts-already-in-db.md) |
| rename views in numerical order | Composer | `RenameViewsSequentially` | [rename-views-in-numerical-order](./references/rename-views-in-numerical-order.md) |
| meta partnos | Composer | `setMetaPartNoFromPartno_cfg` | [meta-partnos](./references/meta-partnos.md) |
| meta descriptions | Composer | `SetMetaDescriptionfromDescription_cfg` | [meta-descriptions](./references/meta-descriptions.md) |
| align pivot | Composer | `AlignPivotToFrame` | [align-pivot](./references/align-pivot.md) |
| set local pivot | Composer | `SetPivotsLocalCoordSystem` | [set-local-pivot](./references/set-local-pivot.md) |
| auto explode | Composer | `ExplodeFirstLevelSubAssembliesActiveView` | [auto-explode](./references/auto-explode.md) |
| Identify Root | Composer | `IdentifyRoot` | [identify-root](./references/identify-root.md) |
| composer connect diagnostic | Composer | `DiagComposer` | [composer-connect-diagnostic](./references/composer-connect-diagnostic.md) |

Choose the views macro by what the view set needs:

- Include bounding-box dimensions on the root: `views [assemblies + all nonnumeric parts + bounding box dims]`. It needs the BestFit proxy actors from `sldworks bounding box` (see that reference).
- Assemblies and every nonnumeric part: `views [assemblies + all nonnumeric parts]`.
- Assemblies and only top-level nonnumeric parts: `views [assemblies + top-level nonnumeric parts]`.
- Only nonnumeric parts: `views [parts-only nonnumeric]`.
- Only specific part numbers from a text file: `views [create balance parts not in db]`.
- Remove views named in a pasted list: `views [delete parts already in db]`. Despite its name, it does not query a database.

## Shared conventions

- Never add an `Attribute VB_Name = ...` line to macro source.
- For Composer view-creation macros using `ZoomFitAll`, save `CameraSmoothTime`,
  set it to `0` before the first zoom, and restore it on completion,
  cancellation, and errors. Report restore failures.
- Logs are written to fixed paths, normally `...\Skyfacture\3 - Economics\skycode\logs\` on OneDrive, or `C:\skycode\logs\`. Each reference lists its log. Check the log after every run; it records what was created, changed, skipped, or failed.
- `_cfg` properties (`Meta.PARTNO_cfg`, `Meta.PartNo_cfg`, `Meta.Description_cfg`) are the configured values. Macros prefer them over the plain `Meta.*` values.
- Numeric part numbers are skipped by the views macros. Nonnumeric and blank values are included.
- Part actors are identified by an actor-path segment that begins with `P`. Assemblies begin with `A`.
- Boolean flags from SOLIDWORKS are normalised with `CBool(value <> False)`, because a native true can be `1` rather than VBA `True` (`-1`).
- The kitify macro never saves the source assembly, and it restores the SOLIDWORKS prompt preferences it changes.
- Composer macros do not save the document. They change names, properties, visibility, or views in memory, so save the document yourself if the change must persist.

## Suggested order for a maintenance cycle

This order follows the data dependencies between macros. Confirm it against the current run before relying on it.

1. Run **kitify and bounding box** (`KitifyAndBoundingBox`) so kitify finishes before the box is added to `Default`. The box is then suppressed in every kit configuration. The 2026-11-19 Composer import did not see that box in the kit configurations; it remains resolved in `Default`. Use the merged BOM workflow in [solidworks-pdm-bom-export](../solidworks-pdm-bom-export/SKILL.md) for the BOM.
2. Import the combined root into Composer, with custom properties and the naming settings described in [kitify](./references/kitify.md).
3. Run **meta partnos** and **meta descriptions** to reconcile `Meta.*` values with their `_cfg` values.
4. Run **align pivot** or **set local pivot** if pivots need correcting, and **auto explode** if an exploded view is needed.
5. Run **bounding box dims** on the Composer document. Use **sldworks bounding box** separately before importing only when another model needs a proxy.
6. Run the **views** macro that matches the view set you need. Use **rename views in numerical order** if the view names must be sequential.
7. Use **views [delete parts already in db]** only with a reviewed list, because deletion is not undone by the macro.
8. Use **Identify Root** and **composer connect diagnostic** only when diagnosing a problem.

## Caveats

- **Identify Root** has hardcoded part numbers and IDs and writes to `C:\Temp\ComposerDebug.txt`.
- **views [create balance parts not in db]** also hardcodes a set of part numbers when it identifies the root assembly.
- **views [delete parts already in db]** deletes views without a restore path.
- **rename views in numerical order** uses the order returned by the Composer API, not a numeric sort.
- Several macros hide collaborations or change visibility and zoom without restoring the previous state.
