# bounding box dims

Macro: [Skycode - bounding box dims.bas](<../../../vba scripts/Skycode - bounding box dims.bas>)

## Purpose

Adds length, width, and height dimensions to the visible bounding-box proxy actors in the Composer document, using their stored origin and edge vectors.

## Host and entry point

- Composer VBA through an Access host that references `composer.exe`. Run `BoundingBoxDimensions`.

## Inputs

- The active Composer document. Proxy actors are named `BoundingBox` or the legacy `BestFitEnvelope`.
- Each proxy needs `Meta.BestFitOrigin` and `Meta.BestFitEdge1` to `Meta.BestFitEdge3`, written by [sldworks-bounding-box](./sldworks-bounding-box.md).

## Outputs

- Three named dimension collaborations per visible, valid proxy: length, width, and height. Labels are whole numbers and offset 100 units.
- Log: `C:\skycode\logs\boundingboxdims.txt`.

## Side effects

- Sets the proxy's opacity to 100.
- Removes earlier dimensions it created and replaces them. The removed dimensions are not restored.
- Saves and restores the document selection.

## Gotchas

- The longest, middle, and shortest edges are labelled length, width, and height.
- Dimensions it created are tagged `SkycodeBBoxDim` and excluded from the scan.
- Hidden proxies have their opacity set but are not dimensioned.
- Proxy name matching is case-insensitive.
