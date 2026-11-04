# Proposed RDX5 SolidWorks Assembly Hierarchy

Proposed kit assemblies use inline-code formatting for readability. Existing assembly
components and other visible tree items are shown in regular text. Kit
assemblies are placed under the corresponding main drifter subassembly; kits
whose current CSVs use the whole drill as their parent remain at drill level.

7393 0800 45-kitifed (Default)  
├─ BG 009 287 40-1-kitifed<1> (RDX5 GH ASSY)  
│  └─ `Gear Housing Kit.SLDASM [proposed; no kit CSV/PN]`  
│     ├─ `7372 1044 03 — RDX5 GH SEAL KIT.SLDASM`  
│     └─ `7372 2822 01 — RDX5 Rotation Kit.SLDASM`  
├─ BG 009 287 40-2-kitifed<2> (RDX5 IMPACT BODY ASSY)  
│  └─ `Impact Body Kit.SLDASM [proposed; no kit CSV/PN]`  
│     ├─ `7372 1045 03 — RDX5 IB SEAL KIT.SLDASM`  
│     └─ `7372 2821 01 — RDX5 Percussion Kit.SLDASM`  
│        ├─ Piston  
│        ├─ Piston seals  
│        └─ Other percussion-kit contents  
├─ BG 009 205 35-kitifed<1> (RDX5 FRONT HEAD ASSY, 45)  
│  └─ `Front Head Kit.SLDASM [proposed; no kit CSV/PN]`  
│     ├─ `7372 1043 01 — HLX-FH-SEAL KIT.SLDASM [in composite BOM; no standalone CSV]`  
│     └─ `7372 1346 03 — RDX5 FH HYDR FITTING KIT.SLDASM`  
├─ `7372 1343 03 — RDX5 DRILL FASTENER KIT.SLDASM [drill-wide]`  
├─ `7372 1347 03 — RDX5 DRILL HYDR FITTING KIT.SLDASM [drill-wide]`  
├─ `7372 1348 03 — RDX5 DRILL BEARING KIT.SLDASM [drill-wide]`  
├─ 280 678 90-kitifed<1> (HOSE CLAMP)  
├─ 7372 1414 01-kitifed<1> (DRILL COLLAR)  
└─ Mates

The Complete Drifter Kit is intentionally not shown as another hardware-owning
assembly node. Model it as a derived configuration/envelope from the governed
subassembly kits, either in this main assembly or in a separate derived
envelope assembly.

This is the proposed kit placement, not a claim that all listed kit assemblies
already exist in the current model. When implementing it, avoid duplicate
component instances: parts represented by a kit assembly should not also
remain active as a second occurrence in the same configuration/BOM.
