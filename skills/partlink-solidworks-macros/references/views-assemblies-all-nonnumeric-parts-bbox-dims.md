# views [assemblies + all nonnumeric parts + bounding box dims]

Macro: [Skycode - views [assemblies + all nonnumeric parts + bounding box dims].bas](<../../../vba scripts/Skycode - views [assemblies + all nonnumeric parts + bounding box dims].bas>)

## Purpose

Creates views for the root assembly, its subassemblies, and eligible nonnumeric parts. Then creates a root bounding-box view (`<root name> BOX`) with three dimensions.

## Host and entry point

- Composer VBA through an Access host. Run `CreateViewsAllWithBoundingBoxDims`.

## Inputs

- The active Composer document and its actors and properties. No prompts or database.
- If the document has several configurations, `ADefault` is used for shared view names.
- Visible proxy actors with `Meta.BestFitOrigin` and `Meta.BestFitEdge1` to `3`. [KitifyAndBoundingBox](./kitifyandboundingbox.md) creates the combined-root proxy, leaves it resolved in `Default`, and suppresses it in kit configurations. The 2026-11-19 Composer import did not see that box in the kit configurations. Otherwise create one first with [sldworks-bounding-box](./sldworks-bounding-box.md), then import it into Composer.

## Outputs

- Root view named from `Meta.PartNo`, or from `Actor.Name` if that is blank.
- Subassembly views named from `Meta.PartNo`, or `SubAssy_<path>`.
- Part views named from the configured part number, or `Part_<path>`.
- The BOX view, with three rounded dimensions.
- Log: `C:\skycode\logs\views [assys + all nonnumeric parts + bounding box dims].txt`.

## Side effects

- Sets `CameraSmoothTime` to 0 and restores its previous value at cleanup.
- Hides collaborations and the ground grid while building views. Prior visibility is not restored.
- Shows message boxes on failures.

## Gotchas

- Only actor IDs that begin with `P` are treated as parts.
- Numeric `Meta.PartNo` values are skipped. Blank and nonnumeric values are included.
- Actors whose names contain `BoundingBox` are treated as proxies for the BOX view. The legacy `BestFitEnvelope` name alone is not matched.
- Dimensions are tagged `SkycodeBBoxDim`, and proxy opacity is set to 100.
