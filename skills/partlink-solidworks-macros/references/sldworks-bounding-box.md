# sldworks bounding box

Macro: [Skycode - sldworks bounding box.bas](<../../../vba scripts/Skycode - sldworks bounding box.bas>)

## Purpose

Creates or validates a native SOLIDWORKS bounding-box feature in best-fit envelope mode, constrained normal to the Front, Top and Right planes (`swBoundingBoxType_CustomPlane`, Front Plane as the alignment entity). It is not a certified global minimum.

## Host and entry point

- SOLIDWORKS VBA, or an Access host with the SOLIDWORKS type library. Run `SldworksBoundingBox` on one saved model, with the intended configuration active. Kitify does not call this routine. The canonical kitified-root path is [kitify and bounding box](./kitifyandboundingbox.md), which calls `GetCurrentModelBoundingBox(False)` only after every kit configuration exists and `Default` is active.

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
- `GetCurrentModelBoundingBox` does not save the host assembly. `KitifyAndBoundingBox` saves only the generated combined assembly after confinement. It does not save the original source.
- `InsertEnvelope` resolves the new envelope in every existing configuration. Deleting it from one configuration deletes it from the others, including `Default`. The combined macro suppresses the envelope and `Skycode_BoundingBox` in every non-Default configuration instead. `SetSuppression2(0)` returns `2` on success. On the 2026-11-19 Composer import of `7397 0500 28`, Composer did not see the box in the kit configurations.

## Gotchas

- The plane is selected with `SelectByID2` (PLANE, then DATUMPLANE), falling back to `Feature.Select2`. Each attempt is logged. The alignment entity is assigned as `PlanarEntity` without `Set`, because it is a Variant property.
- An existing root bounding-box feature is deleted and recreated on each run. More than one root bounding-box feature raises an error.
- The sketch must contain 12 linear edges in three groups of four.
- Dimensions are reported longest to shortest, not along world X, Y, or Z.
- Replacing an existing proxy requires that the file already contains the named generated feature.
- The Composer-side macros read these proxies through `Meta.BestFitOrigin` and `Meta.BestFitEdge1` to `3`. See [bounding-box-dims](./bounding-box-dims.md).
