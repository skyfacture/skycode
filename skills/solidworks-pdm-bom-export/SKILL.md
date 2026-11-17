---
name: solidworks-pdm-bom-export
description: 'Export SOLIDWORKS PDM computed BOMs for every assembly in a vault folder to CSV from PowerShell, using the EdmInterface typelib, then merge kit BOM CSVs into the active drifter BOM. Use when asked to retrieve PDM BOM CSVs, export a vault folder BOM, merge kit BOMs into a drifter BOM, or reproduce the Partlink CSV export without the docs/solidworksapi reference.'
---

# Export PDM computed BOMs to CSV

Drive `ConisioLib.EdmVault` from PowerShell and reach the computed-BOM API through typed interop casts generated from `EdmInterface.dll`. This is the same vault pattern the kitify sync uses (see `docs/kitify sync.md`, the PDM access section, and `vba scripts/Skycode - kitify.bas`, the `GetFirstFilePosition`/`GetNextFile` loop around lines 796-840).

Source of truth: the typelib at `C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS PDM\EdmInterface.dll`. The SOLIDWORKS API help page at `https://help.solidworks.com/2025/english/Api/sldworksapiprogguide/Welcome.htm` is client-rendered and returns 403 to `web_fetch`, so use the typelib instead.

## Procedure

1. Log in and resolve the folder:

   ```powershell
   $vault = New-Object -ComObject ConisioLib.EdmVault
   $vaultName = $vault.GetVaultNameFromPath($KitsPath)
   $vault.Login($env:PDM_USER, $env:PDM_PWD, $vaultName)
   $folder = $vault.GetFolderFromPath($KitsPath)
   ```

   Run the export in a child `powershell -File` process with `WaitForExit` and a timeout. Pass credentials through environment variables set by the parent and removed afterward. Never write them to disk. `LoginAuto` hung in this environment. The parent's `Login` call prompted for credentials instead.

   `GetFolderFromPath` returns `$null` if the path is not in the vault. Stop and report that.

2. Enumerate files. End the loop on HRESULT `-2147220984` (`PDM_END_OF_LIST`):

   ```powershell
   $pos = $folder.GetFirstFilePosition()
   while ($true) {
       try { $f = $folder.GetNextFile($pos) }
       catch { if ($_.Exception.HResult -eq -2147220984) { break }; throw }
       # filter on extension, e.g. .sldasm
   }
   ```

3. Build the interop assembly. Load the typelib with `LoadTypeLibEx` (P/Invoke, `oleaut32.dll`, `CharSet.Unicode`), wrap the `IntPtr` out-parameter with `Marshal.GetObjectForIUnknown`, then call `TypeLibConverter.ConvertTypeLibToAssembly` with namespace `EdmInterop`. The notify-sink argument must be a non-null `ITypeLibImporterNotifySink` (not the exporter variant). Using an `out object` parameter fails with "Specified OLE variant is invalid". The returned `AssemblyBuilder` is not written to disk. Call `.Save(name)` from a `%TEMP%` working directory, then move the file.

4. Cast PDM objects to the interface that owns the method. Late-bound calls fail for methods on newer interfaces:

   ```powershell
   function Typed($obj, [string]$iface) {
       if ($obj -is [System.Management.Automation.PSObject]) { $obj = $obj.BaseObject }
       [Runtime.InteropServices.Marshal]::GetTypedObjectForIUnknown(
           [Runtime.InteropServices.Marshal]::GetIUnknownForObject($obj),
           $script:asm.GetType('EdmInterop.' + $iface))
   }
   ```

   Call with `MethodInfo.Invoke`. Unwrap `TargetInvocationException` to get the real error. Unwrap `PSObject` arguments first, or reflection fails with "Object of type 'PSObject' cannot be converted to type 'String'".

5. Compute and export each BOM:

   ```powershell
   $view = Call $file 'IEdmFile7' 'GetComputedBOM' @([object]1, [int]0, 'Default', [int]3)
   Call $view 'IEdmBomView3' 'SaveToCSV' @([string]$rawPath, [bool]$true)
   ```

   Layout `1` is "BOM" (computed). `2` is "Weldment Cut List" and `3` is "Weldment BOM". Config `'Default'` worked on the Partlink kit assemblies. If it throws, fall back to `'@'`. `SaveToCSV`'s second argument is the Level-column flag, so pass `$true`.

   The `flags` argument is `EdmBomFlag`: `EdmBf_AsBuilt = 1`, `EdmBf_ShowSelected = 2`. Use `3` (both). `3` writes the root row, the same as the PDM UI's "Show Selected" option, and keeps the as-built children. `1` drops the root row. `2` alone expands nested children differently from the as-built tree, so it changes the child rows.

6. Reshape the native output to the repo CSV convention (see Output conventions), then write it.

## API reference (from the typelib)

| Interface | Member | Signature / notes |
|---|---|---|
| `IEdmFile7` | `GetComputedBOM` | `(Object layout, Int32 version, String config, Int32 flags)` returns `IEdmBomView`. It is not on `IEdmFile5`. |
| `IEdmBomView3` | `SaveToCSV` | `(String path, Boolean bCreateLevelColumn)`. Writes UTF-16 LE with BOM and no CRLF. The root row is included when the BOM was computed with flag `3`. |
| `IEdmBomView` | `GetRows` | `(out Object[])` returns the rows. With flag `3`, row 0 is the root. |
| `IEdmBomCell` | `GetVar` | `(lVarID, EdmBomColumnType, out value, out computed, out config, out readOnly)` |
| `IEdmBomCell` | `GetTreeLevel` | Returns `Int32`. |
| `IEdmFile5` | `CurrentState`, `CurrentRevision`, `CurrentVersion`, `ID`, `Name`, `Vault`, `GetLocalPath(Int32)` | File metadata for the root row. Variables such as Description, Material, and Weight come from `GetEnumeratorVariable`. |

Errors seen during this work:

- HRESULT `0x8004020B` "Not all the objects requested could be removed" came from a null string argument.
- "Specified OLE variant is invalid" came from an `out object` parameter. Use `IntPtr`.

## Output conventions (repo CSV format)

Match the 18-column reference `partlink cad sandbox\kitified-generated_*\7397 0320 22.csv`:

- Header: `Level,File Name,Part Number,Configuration,Description,Qty,State,Revision,Version,Date Modified,ID,G.P.N,NOTE,OEM,OLD PartNo,Found In,Material,Weight`
- Root row first, with `Level` `1` and the assembly's own metadata. With flag `3`, the native export supplies it, including Description (variable 57) and Date Modified (`EdmBomCol_Date`). Keep the native values.
- Child levels are dotted under the root. The native export numbers direct children `1, 2, 3`, so prefix every child level with `1.` to get `1.1` (direct child) and `1.1.1` (grandchild).
- Encoding is UTF-8 with BOM and CRLF line endings, with comma delimiters. Quote only fields containing `,`, `"`, CR, or LF, and double inner quotes.
- Child `File Name` values from the native export carry leading indentation. Keep it to match the reference.

## Known gaps

- Earlier, the root row's Description and Date Modified were blank because the BOM was computed with flag `1`. Flag `3` fills both. Root Date Modified matched the root working copy's local timestamp, and a child copy's timestamp matched its BOM date, so PDM stamps local files with the version date.
- Native export dates are `dd/MM/yyyy`. The reference's date convention was not confirmed for these vault dates, so check before relying on date order.

## Validation

- Every file starts with the UTF-8 BOM (`EF BB BF`) and has no bare LF.
- The header equals the reference header exactly.
- The first row has `Level` `1`, and all later levels start with `1.`.
- Root row (`Level` `1`) has a non-empty Description and Date Modified.
- Child rows equal the flag-`1` children from the as-built export. Flag `3` must add only the root row.
- Run this on every assembly, not a sample.

## Merge kit BOMs into the drifter BOM

Run this after the kit CSVs have been exported from the active drifter assembly. The kit CSVs are the per-kit BOMs generated from that assembly, and they are appended to its BOM as additional rows. They come from the same export run, in the `bom-export-12-KITS_*` folder next to the base. Do not use `kitify\`: it holds kit CSVs from earlier assemblies, such as RDX5 from 7393 0800 45, and merging those into another assembly's BOM adds unrelated rows.

Inputs:

- **Base (drifter BOM):** the 18-column computed BOM of the active assembly, for example `partlink cad sandbox\kitified-generated_*\7397 0320 22.csv`. Do not substitute a different assembly's BOM. Using `7393 0800 45.csv` as the base for `7397 0320 22` was a mistake, so confirm the base root `File Name` matches the active assembly before merging.
- **Kit BOMs:** every other CSV in the `bom-export-12-KITS_*` folder, for example `7372 2927 02.csv`. Each file name is its kit root's Part Number. Exclude the base file and any previous merged output.

Procedure:

1. Check each kit CSV. The header must equal the 18-column base header, including `Weight`. Kit files start with a UTF-8 BOM (`EF BB BF`) and use CRLF endings, so strip the BOM before reading.
2. Read every file as Latin-1 (`[Text.Encoding]::GetEncoding(28591)`) so the bytes round-trip exactly. Do not decode as UTF-8, which would corrupt characters such as the `�` in `ROTARY SEAL (30x37.5x3.2 Ø30, FKM)`.
3. Start with the base file's text unchanged. Add a CRLF if it does not end with one.
4. For each kit file, skip the header line and every empty line. Append each remaining row followed by CRLF. Only a row with 17 fields needs a comma first, which pads it with an empty `Weight`. Rows with 18 fields are copied unchanged.
5. Write the result as bytes in Latin-1. The output is `bom-export-12-KITS_*\<active assembly PartNo>.csv`, for example `bom-export-12-KITS_20261117_152548\7397 0320 22.csv`.
6. Keep the native `Level` values of the kit rows (`1`, `1.1`, and so on) and do not renumber them. Row order does not matter. Kit roots stay at `Level` `1`, so there is one `Level` `1` row per kit plus the assembly root.

Validation:

- Row count equals base data rows plus the sum of kit data rows, counted by `Import-Csv`.
- Count fields with `Import-Csv`, not by splitting on commas. The base has quoted fields that contain commas, such as `"O-RING (85x2, 70SHORE)"`, so a naive split gives false results. Every row must parse to the 18 header columns.
- The number of `Level` `1` rows equals 1 (the assembly root) plus the number of kit CSVs.
- Each kit's root row (`Level` `1`, with the Part Number equal to the kit's file name) is present, and every kit data row appears verbatim in the output.
- The base file's bytes are a prefix of the output.

Remove any earlier merged output that was built on a different base, for example `drifter root partno.csv`, after the new file is verified.

## Cleanup

Delete the temporary raw exports and any interop DLLs written during the run. Keep only the final CSV folder. Do not leave failed attempt folders in the output root.
