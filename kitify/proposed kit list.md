# Proposed RDX5 Kit List

## Proposed SolidWorks kit hierarchy

```text
7393 0800 45.SLDASM
├─ Front Head Kit.SLDASM                         [proposed; no kit CSV/PN]
│  ├─ 7372 1043 01 — HLX-FH-SEAL KIT.SLDASM     [in composite BOM; no standalone CSV]
│  └─ 7372 1346 03 — RDX5 FH HYDR FITTING KIT.SLDASM
├─ Gear Housing Kit.SLDASM                       [proposed; no kit CSV/PN]
│  ├─ 7372 1044 03 — RDX5 GH SEAL KIT.SLDASM
│  └─ 7372 2822 01 — RDX5 Rotation Kit.SLDASM
├─ Impact Body Kit.SLDASM                        [proposed; no kit CSV/PN]
│  ├─ 7372 1045 03 — RDX5 IB SEAL KIT.SLDASM
│  └─ 7372 2821 01 — RDX5 Percussion Kit.SLDASM
├─ 7372 1343 03 — RDX5 DRILL FASTENER KIT.SLDASM     [drill-wide]
├─ 7372 1347 03 — RDX5 DRILL HYDR FITTING KIT.SLDASM [drill-wide]
└─ 7372 1348 03 — RDX5 DRILL BEARING KIT.SLDASM      [drill-wide]
```

The principal kit identities are listed in `kit list.txt`. The Front Head
hydraulic kit is explicitly Front Head-only. The fastener, hydraulic-fitting,
and bearing kit CSVs instead use the whole drill as their parent, so they are
shown at drill level rather than assigned to a specific major subassembly.

The unnumbered Front Head, Gear Housing, and Impact Body kit assemblies are
proposed modelling identities; no part numbers for these parent kits are
established by the supplied CSVs.

## Seal-kit composite hierarchy

The seal-kit CSVs show a separate composite product structure:

```text
7372 1046 03 — RDX5 SEAL KIT COMP
└─ 7372 1392 03 — RDX5-GI SEAL KIT
   ├─ 7372 1044 03 — RDX5 GH SEAL KIT
   ├─ 7372 1045 03 — RDX5 IB SEAL KIT
   └─ 7372 1043 01 — HLX-FH-SEAL KIT
```

This is a valid nesting for the seal-kit product. Do not also insert these
composite wrappers into the three major kits if doing so would duplicate
seal-kit geometry in the same SolidWorks configuration.

## Complete Drifter Kit

The Complete Drifter Kit remains a derived configuration or envelope in the
main drifter assembly (or a separate derived envelope assembly). It is not
another hardware-owning node; the subassembly kits remain the source of truth
for kit composition.
