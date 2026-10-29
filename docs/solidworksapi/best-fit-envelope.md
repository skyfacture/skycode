# SOLIDWORKS native Best Fit envelope

Verified against the installed SOLIDWORKS 2025 API and official 2025 online API
help on 2026-10-29. Consult this directory before fetching additional API help.
It was empty when this investigation began.

## Scope and limitation

The initial implementation uses SOLIDWORKS' native **Best Fit rectangular
bounding box**, as approved for this first implementation. This is not a
mathematically certified global-minimum-volume box. The consulted API reference
does not provide that guarantee. It is also not a convex hull, material volume,
or a packing calculation.

The calculation uses the active part or root assembly and its active
configuration. Hidden solid bodies and hidden components are included.
Surfaces and assembly envelope/reference components are excluded. Suppressed
geometry is outside the active configuration's geometry.

## Working script

The VBA source is now [Skycode - sldworks bounding box.bas](../../Skycode%20-%20sldworks%20bounding%20box.bas),
renamed from the initial extensionless working file.

1. Copy the source into a standard VBA module in Access or a SOLIDWORKS macro.
2. Under **Tools > References**, enable the installed **SOLIDWORKS type library**.
   The installed file used during investigation was
   `C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\sldworks.tlb`.
   A separate constants-library reference is not required.
3. Keep the intended part or assembly active in the running SOLIDWORKS session.
4. Run `SldworksBoundingBox`, or call `GetCurrentModelBestFitEnvelope`.
   `GetCurrentModelBestFitEnvelope(False)` returns the report without a message
   box, for programmatic use. Both forms first show a **Save As** dialog for
   the TXT report. Cancelling returns an empty string without changing the
   model or creating output files.

The script attaches using `GetObject(, "SldWorks.Application")`, rather than
launching another SOLIDWORKS instance. It gets `ActiveDoc` and rejects missing
models and drawings.

A new feature is named `Skycode_BestFitEnvelope`. An existing root bounding-box
feature is reused only when its settings match the requested geometry scope.
Conflicting settings and multiple root boxes produce explicit errors rather
than silently modifying an existing feature.

Creating the feature changes the open model and clears the current selection.
The script rebuilds the model but does **not** save the source model or write
custom properties. It saves a separate solid proxy part beside the saved source
model and a text report at the location chosen in the Save As dialog.
For assemblies, the proxy is inserted as a fixed
envelope component and excluded from the BOM; tree ordering is not changed.
Save the source assembly manually to persist that component reference.
On failure the script logs the stage and raises the error; partially created
files or components may remain and are explicitly reported.

## Verified API path

- `IFeatureManager.CreateDefinition(swFmBoundingBox)` returns
  `IBoundingBoxFeatureData`. The installed constants library gives
  `swFmBoundingBox = 114`.
- `ReferenceFaceOrPlane = swBoundingBoxType_BestFit`, whose installed value is
  `1`, requests the native Best Fit mode.
- Set `IncludeHiddenBodies = True` and `IncludeSurfaces = False`.
- For assemblies also set `IncludeHiddenComponents = True` and
  `IncludeEnvelopeComponents = False`.
- `IFeatureManager.CreateFeature` performs the native calculation and returns
  the feature.
- On the tested assembly the feature's `GetTypeName2` was
  `BoundingBoxProfileFeat`; `GetSpecificFeature2` returned an `ISketch`.
- `ISketch.GetSketchSegments` returned 12 straight edges. Each dimension appears
  four times. Sorting all 12 edge lengths and validating three groups of four
  handles rectangular boxes, square prisms and cubes without discarding equal
  dimensions.
- API lengths are in metres. Convert to millimetres using `* 1000`.
  Box volume is the product of the three lengths in metres; litres are
  cubic metres multiplied by 1000.
- Report dimensions longest to shortest; they are box-edge dimensions, not
  world-axis extents.
- The tested COM server returned noncanonical true values to VBA: a property
  displayed `True` while `Not property` also evaluated as `True`. Compare
  returned Boolean values to `False` explicitly rather than using `Not` or
  comparing to `True`. This applies to the inclusion flags and rebuild result.

`IAssemblyDoc.GetBox` is not used: its approximate world-axis box does not
implement this Best Fit requirement.

`IFeatureManager.InsertGlobalBoundingBox` is documented as obsolete. Its
replacement is `CreateDefinition` with `IBoundingBoxFeatureData`, followed by
`CreateFeature`. `IBoundingBoxFeatureData` itself exposes fit and inclusion
settings, not three dimension properties.

## Evidence and sources

The installed API help was found at
`C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\sldworksapi.chm`.
Windows Help extraction attempts produced no files. The online pages below
were then read using the browser because direct HTTP fetches returned 403.
These notes summarize the findings rather than copying the vendor reference.

- [IBoundingBoxFeatureData members](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IBoundingBoxFeatureData_members.html)
- [ReferenceFaceOrPlane](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IBoundingBoxFeatureData~ReferenceFaceOrPlane.html)
- [Bounding-box fit enumeration](https://help.solidworks.com/2025/english/api/swconst/SolidWorks.Interop.swconst~SolidWorks.Interop.swconst.swGlobalBoundingBoxFitOptions_e.html)
- [Feature identifier enumeration](https://help.solidworks.com/2025/english/api/swconst/SolidWorks.Interop.swconst~SolidWorks.Interop.swconst.swFeatureNameID_e.html)
- [Creating a bounding box, VBA example](https://help.solidworks.com/2025/english/api/sldworksapi/Create_Bounding_Box_Example_VB.htm)
- [GetSpecificFeature2](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature~GetSpecificFeature2.html)
- [Obsolete InsertGlobalBoundingBox](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeatureManager~InsertGlobalBoundingBox.html)

## Validation status

A typed COM probe connected to the running SOLIDWORKS revision `33.5.0`,
obtained the active assembly, and created the native Best Fit feature with the
settings above. Its sketch returned 12 line segments in three groups of four.
No save was requested.

The actual working VBA source was imported into a new, isolated Access
database with the installed SOLIDWORKS type-library reference. It compiled
and ran successfully twice against `RDX5 Complete Kit`, configuration
`Default`, reusing the existing feature without creating a duplicate.
Both calls returned the same report:

- Length: 977.982 mm
- Width: 389.232 mm
- Height: 248.906 mm
- Box volume: 0.094749096 m^3 (94.749096 litres)

The initial feature-creation path was verified using the typed COM probe;
the VBA runtime test exercised the existing-feature path. Neither saved the
SOLIDWORKS model. Only the currently active configuration was measured;
other configurations, part documents, and specialized assembly modes need
separate runtime checks.

## Solid proxy geometry for Composer

The current feature is a bounding-box sketch, not a solid body. The local
Composer command-line help documents a curve-import setting,
`Batch.IOImportCurves`, but does not establish that the special SOLIDWORKS
`BoundingBoxProfileFeat` sketch is imported. Do not promise that merely
showing this feature will create a Composer geometry actor.

The script now creates a separate part with an actual solid body from the
native box's corners and edge directions, positioned in source model
coordinates. It converts sketch coordinates through the inverse
`ModelToSketchTransform`, identifies the three incident edges at a corner,
constructs a right-handed placement transform and applies it to a modeler box.
`CreateFeatureFromBody3` stores the resulting body as part geometry.

The source model must already have a saved path, and SOLIDWORKS must have a
valid default part template. Outputs are named
`<source-name>_BestFitEnvelope.SLDPRT` in the source model's folder, not the
VBA script's folder. The TXT report's Save As dialog suggests a matching
filename and initially opens in the source folder; choose its name and location
before processing begins. The dialog supplies `.txt` when no extension is
entered and confirms replacement of an existing report. It uses the Windows
common dialog API, compatible with 32-bit and 64-bit VBA hosts.
The proxy is overwritten on subsequent runs. The existing generated body feature
is replaced in the same part document and saved, including when that document
is loaded through the source assembly. An existing output lacking the named
generated feature is rejected rather than replacing an unrelated part.
The assembly reuses a top-level component referencing that exact proxy path
instead of inserting another instance. Older timestamped outputs/components
are not automatically removed.
The text file contains the same report string displayed in the popup and is
written as Unicode text before the popup appears.
Reported length, width and height are rounded to the nearest whole millimetre.
The proxy geometry and volume calculation retain full precision.

For assemblies `InsertEnvelope` inserts the saved part, identity placement
preserves the already-transformed solid coordinates, and the component is
fixed and excluded from the BOM. No insertion occurs for a source part model.
The new proxy part remains open. The source model is reactivated and remains
unsaved. Enable envelope-component import in Composer and save the assembly
before importing its on-disk contents. Transfer into Composer has not yet
been verified; confirm actor inclusion and alignment there.

The implementation was compiled and executed from an isolated Access database
against the active assembly. Validation confirmed one solid body, eight
vertices matching every source sketch endpoint within approximately
`5.6e-17 m`, volume `0.094749096433889 m^3`, fixed/envelope/BOM-exclusion flags,
and an existing text report. The initial failed-test component was removed;
the successful proxy component remains in the unsaved assembly.

The three sorted lengths currently reported are insufficient to reconstruct
the original oriented box: its position and orientation must also be retained.
If the proxy becomes an assembly component, keep it out of subsequent envelope
calculations (for example, make it an envelope component and retain
`IncludeEnvelopeComponents = False`). Adding an ordinary solid box component
can also affect assembly mass properties and BOMs; do not treat it as a
production component.

An alternative is to create the box directly in Composer. Its documented
`CreateBBox` method accepts two diagonal corner points but no orientation
argument. An arbitrarily oriented Best Fit box would therefore require a
separate, verified placement/rotation step; two opposite world-space corners
alone do not describe the full orientation.

Local references:

- [Composer import settings sample](../swcomposerapi_html/APIFiles/CommandLineParameters.htm)
- [Composer CreateBBox reference](../swcomposerapi_html/APIFiles/API_AutoGenerate/group__group_composer_a_p_i_document_method.htm)

Additional official API references consulted:

- [CreateBodyFromBox3](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModeler~CreateBodyFromBox3.html)
- [InsertEnvelope](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IAssemblyDoc~InsertEnvelope.html)

`CreateBodyFromBox3` takes nine doubles: one face centre, the box axis,
width, length and height. The implementation first builds a local box with
axis `(0,0,1)` and then applies the full orientation transform; specifying
the axis alone would not preserve rotation about that axis.
For VBA the component transform uses
`component.Transform2 = transform` (without `Set`); a `Set` assignment failed
at runtime with error 438 on the installed API.

The fixed-filename overwrite revision compiled in Access. After an initial
connection failure (VBA error 429), a live retest ran the actual script twice
successfully against `RDX5 Complete Kit`, configuration `Default`. Both calls
used identical output paths and returned identical reports. A sentinel appended
to the report between runs was removed by the second run, and the final text
matched the popup report exactly. Exactly one top-level component referenced
the fixed proxy path. The overwritten part contained one solid body with eight
vertices; its volume matched the current sketch at `0.0947491544624082 m^3`.
Fixed, envelope and BOM-exclusion flags were verified. The source assembly
remained unsaved. Older timestamped proxy components were not removed.
