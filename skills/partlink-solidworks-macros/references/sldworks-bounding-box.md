# sldworks bounding box

Macro: [Skycode - sldworks bounding box.bas](<../../../vba scripts/Skycode - sldworks bounding box.bas>)

## Purpose

Creates or validates a native SOLIDWORKS bounding-box feature in best-fit envelope mode, constrained normal to the Front, Top and Right planes (`swBoundingBoxType_CustomPlane`, Front Plane as the alignment entity). It is not a certified global minimum.

## Host and entry point

- SOLIDWORKS VBA, or an Access host with the SOLIDWORKS type library. Run `SldworksBoundingBox`, which calls `GetCurrentModelBoundingBox`. Kitify calls that function with messages disabled; both modules must be in the same VBA project.

## Inputs

- An open SOLIDWORKS part or assembly that is saved to disk. The active configuration is used.
- The `.txt` report is written beside the source model and overwritten without a dialog.
- For assemblies, hidden components are included and envelope components are excluded.
- The `_cfg` properties are not read.

## Outputs

- A sibling proxy part `<source>_BoundingBox.SLDPRT`. Its custom properties are `BestFitOrigin` and `BestFitEdge1` to `BestFitEdge3`, in millimetres in model coordinates.
- A text report with the configuration, dimensions from longest to shortest, volume, and paths.
- For an assembly, the proxy is inserted as a fixed envelope component and excluded from the BOM.

## Side effects

- Adds a bounding-box feature to the source model in memory. Any existing `BoundingBoxProfileFeat` is deleted and recreated. The source is not saved by the macro. An assembly may be modified by the inserted envelope.
- No rollback. A failure partway through can leave a new feature or component behind.
- When called by Kitify, the model being modified and subsequently saved is the combined output, not the original source assembly.

## Gotchas

- The plane is selected with `SelectByID2` (PLANE, then DATUMPLANE), falling back to `Feature.Select2`. Each attempt is logged. The alignment entity is assigned as `PlanarEntity` without `Set`, because it is a Variant property.
- An existing root bounding-box feature is deleted and recreated on each run. More than one root bounding-box feature raises an error.
- The sketch must contain 12 linear edges in three groups of four.
- Dimensions are reported longest to shortest, not along world X, Y, or Z.
- Replacing an existing proxy requires that the file already contains the named generated feature.
- The Composer-side macros read these proxies through `Meta.BestFitOrigin` and `Meta.BestFitEdge1` to `3`. See [bounding-box-dims](./bounding-box-dims.md).
