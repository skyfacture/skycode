<#
.SYNOPSIS
    Set Razer Joro static colour and brightness over Bluetooth LE, without Synapse.

.EXAMPLE
    .\joro-rgb.ps1 -Color FF0000
    .\joro-rgb.ps1 -Color 00FFFF -Brightness 40
    .\joro-rgb.ps1 -Brightness 100

.NOTES
    Protocol (observed from Synapse Bluetooth traffic):
      Service 52401523-f97c-7f90-0e7f-6c6f4e36db1c
      Write   52401524-...  each command = 8-byte header write, then a separate data write
      Notify  52401525-...  replies
      Header: [txid][dataLen][00][00][class][cmd][01][00]
      Static colour : class 10 cmd 03, data 01 00 00 01 RR GG BB
      Brightness    : class 10 cmd 05, data <0-255>
#>
[CmdletBinding()]
param(
    [ValidatePattern('^#?[0-9A-Fa-f]{6}$')]
    [string]$Color,

    [ValidateRange(0, 100)]
    [int]$Brightness = -1,

    # Bluetooth address, e.g. FAA3F314A655. Auto-detected if omitted.
    [string]$Address
)

$ErrorActionPreference = 'Stop'
if (-not $Color -and $Brightness -lt 0) { throw 'Specify -Color and/or -Brightness.' }

$ServiceUuid = [guid]'52401523-f97c-7f90-0e7f-6c6f4e36db1c'
$WriteUuid   = [guid]'52401524-f97c-7f90-0e7f-6c6f4e36db1c'
$NotifyUuid  = [guid]'52401525-f97c-7f90-0e7f-6c6f4e36db1c'

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Devices.Bluetooth.BluetoothLEDevice, Windows.Devices.Bluetooth, ContentType = WindowsRuntime]
$null = [Windows.Devices.Bluetooth.GenericAttributeProfile.GattDeviceServicesResult, Windows.Devices.Bluetooth, ContentType = WindowsRuntime]
$null = [Windows.Devices.Bluetooth.GenericAttributeProfile.GattCharacteristicsResult, Windows.Devices.Bluetooth, ContentType = WindowsRuntime]
$null = [Windows.Devices.Bluetooth.GenericAttributeProfile.GattWriteResult, Windows.Devices.Bluetooth, ContentType = WindowsRuntime]
$null = [Windows.Devices.Bluetooth.GenericAttributeProfile.GattCommunicationStatus, Windows.Devices.Bluetooth, ContentType = WindowsRuntime]

$asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
} | Select-Object -First 1

function Await($operation, [type]$resultType) {
    $task = $asTask.MakeGenericMethod($resultType).Invoke($null, @($operation))
    if (-not $task.Wait(10000)) { throw 'Bluetooth operation timed out.' }
    $task.Result
}

if (-not $Address) {
    $node = Get-PnpDevice -PresentOnly -Class Bluetooth |
        Where-Object { $_.InstanceId -match '^BTHLE\\DEV_[0-9A-F]{12}' -and $_.FriendlyName -match 'Joro' } |
        Select-Object -First 1
    if (-not $node) { throw 'Joro not found. Make sure it is paired and connected over Bluetooth.' }
    $Address = [regex]::Match($node.InstanceId, 'DEV_([0-9A-F]{12})').Groups[1].Value
}
$addressValue = [Convert]::ToUInt64(($Address -replace '[:-]', ''), 16)

$device = Await ([Windows.Devices.Bluetooth.BluetoothLEDevice]::FromBluetoothAddressAsync($addressValue)) ([Windows.Devices.Bluetooth.BluetoothLEDevice])
if (-not $device) { throw "Could not open Joro at $Address." }

$service = $null
try {
    $services = Await ($device.GetGattServicesForUuidAsync($ServiceUuid)) ([Windows.Devices.Bluetooth.GenericAttributeProfile.GattDeviceServicesResult])
    if ($services.Status -ne 'Success' -or $services.Services.Count -eq 0) { throw "Razer lighting service unavailable ($($services.Status))." }
    $service = $services.Services[0]

    $null = [Windows.Devices.Bluetooth.GenericAttributeProfile.GattOpenStatus, Windows.Devices.Bluetooth, ContentType = WindowsRuntime]
    $open = Await ($service.OpenAsync([Windows.Devices.Bluetooth.GenericAttributeProfile.GattSharingMode]::SharedReadAndWrite)) ([Windows.Devices.Bluetooth.GenericAttributeProfile.GattOpenStatus])
    if ($open -ne 'Success' -and $open -ne 'AlreadyOpened') {
        throw "Lighting service is locked ($open). Quit Razer Synapse (and Razer services) first, then retry."
    }

    $chars = Await ($service.GetCharacteristicsAsync()) ([Windows.Devices.Bluetooth.GenericAttributeProfile.GattCharacteristicsResult])
    if ($chars.Status -ne 'Success') { throw "Could not read Razer characteristics ($($chars.Status)). Quit Razer Synapse first, then retry." }
    $writeChar  = $chars.Characteristics | Where-Object { $_.Uuid -eq $WriteUuid } | Select-Object -First 1
    $notifyChar = $chars.Characteristics | Where-Object { $_.Uuid -eq $NotifyUuid } | Select-Object -First 1
    if (-not $writeChar) { throw 'Razer command characteristic not found.' }

    # Synapse enables reply notifications before sending commands.
    if ($notifyChar) {
        $null = Await ($notifyChar.WriteClientCharacteristicConfigurationDescriptorAsync(
            [Windows.Devices.Bluetooth.GenericAttributeProfile.GattClientCharacteristicConfigurationDescriptorValue]::Notify)) ([Windows.Devices.Bluetooth.GenericAttributeProfile.GattCommunicationStatus])
    }

    $script:txid = Get-Random -Minimum 0x20 -Maximum 0xF0

    function Write-Packet([byte[]]$bytes) {
        $buffer = [System.Runtime.InteropServices.WindowsRuntime.WindowsRuntimeBufferExtensions]::AsBuffer($bytes)
        $result = Await ($writeChar.WriteValueWithResultAsync($buffer, [Windows.Devices.Bluetooth.GenericAttributeProfile.GattWriteOption]::WriteWithResponse)) ([Windows.Devices.Bluetooth.GenericAttributeProfile.GattWriteResult])
        if ($result.Status -ne 'Success') { throw "Write failed: $($result.Status)" }
    }

    function Send-Command([byte]$class, [byte]$cmd, [byte[]]$data) {
        $script:txid = ($script:txid + 1) -band 0xFF
        Write-Packet ([byte[]]@($script:txid, $data.Length, 0, 0, $class, $cmd, 1, 0))
        Write-Packet $data
        Start-Sleep -Milliseconds 50
    }

    if ($Color) {
        $hex = $Color.TrimStart('#').ToUpper()
        $r = [Convert]::ToByte($hex.Substring(0, 2), 16)
        $g = [Convert]::ToByte($hex.Substring(2, 2), 16)
        $b = [Convert]::ToByte($hex.Substring(4, 2), 16)
        Send-Command 0x10 0x03 ([byte[]]@(1, 0, 0, 1, $r, $g, $b))
        Write-Output "Colour set to #$hex"
    }

    if ($Brightness -ge 0) {
        $level = [byte][Math]::Round($Brightness * 255 / 100)
        Send-Command 0x10 0x05 ([byte[]]@($level))
        Write-Output "Brightness set to $Brightness%"
    }
}
finally {
    # Windows PowerShell can't always call Dispose on WinRT objects; release is best-effort.
    foreach ($obj in @($service, $device)) {
        if ($obj) { try { $obj.Dispose() } catch { } }
    }
}
