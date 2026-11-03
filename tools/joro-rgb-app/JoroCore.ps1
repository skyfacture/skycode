# Razer Joro Bluetooth LE lighting protocol (observed from Synapse traffic).
#   Service 52401523-f97c-7f90-0e7f-6c6f4e36db1c
#   Write   52401524-...  each command = 8-byte header write, then a separate data write
#   Notify  52401525-...  replies
#   Header: [txid][dataLen][00][00][class][cmd][01][00]
#   Static colour : class 10 cmd 03, data 01 00 00 01 RR GG BB
#   Brightness    : class 10 cmd 05, data <0-255>

$ErrorActionPreference = 'Stop'

$script:JoroServiceUuid = [guid]'52401523-f97c-7f90-0e7f-6c6f4e36db1c'
$script:JoroWriteUuid   = [guid]'52401524-f97c-7f90-0e7f-6c6f4e36db1c'
$script:JoroNotifyUuid  = [guid]'52401525-f97c-7f90-0e7f-6c6f4e36db1c'
$script:JoroLockMessage = 'Razer Synapse is holding the keyboard. Quit Synapse and try again.'
$script:JoroSleepMessage = 'Keyboard not responding. Press a key to wake it, then retry.'

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Devices.Bluetooth.BluetoothLEDevice, Windows.Devices.Bluetooth, ContentType = WindowsRuntime]
$null = [Windows.Devices.Bluetooth.GenericAttributeProfile.GattDeviceServicesResult, Windows.Devices.Bluetooth, ContentType = WindowsRuntime]
$null = [Windows.Devices.Bluetooth.GenericAttributeProfile.GattCharacteristicsResult, Windows.Devices.Bluetooth, ContentType = WindowsRuntime]
$null = [Windows.Devices.Bluetooth.GenericAttributeProfile.GattWriteResult, Windows.Devices.Bluetooth, ContentType = WindowsRuntime]
$null = [Windows.Devices.Bluetooth.GenericAttributeProfile.GattCommunicationStatus, Windows.Devices.Bluetooth, ContentType = WindowsRuntime]
$null = [Windows.Devices.Bluetooth.GenericAttributeProfile.GattOpenStatus, Windows.Devices.Bluetooth, ContentType = WindowsRuntime]

$script:JoroAsTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
} | Select-Object -First 1

$script:Joro = $null
$script:JoroTxId = Get-Random -Minimum 0x20 -Maximum 0xF0

function Wait-JoroAsync($Operation, [type]$ResultType) {
    $task = $script:JoroAsTask.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    if (-not $task.Wait(10000)) { throw $script:JoroSleepMessage }
    $task.Result
}

function Assert-JoroStatus($Status) {
    switch ([string]$Status) {
        'Success' { return }
        'AccessDenied' { throw $script:JoroLockMessage }
        default { throw $script:JoroSleepMessage }
    }
}

function Get-JoroAddress {
    $node = Get-PnpDevice -PresentOnly -Class Bluetooth -ErrorAction SilentlyContinue |
        Where-Object { $_.InstanceId -match '^BTHLE\\DEV_[0-9A-F]{12}' -and $_.FriendlyName -match 'Joro' } |
        Select-Object -First 1
    if (-not $node) { throw 'No paired Joro found. Pair it in Windows Bluetooth settings.' }
    [Convert]::ToUInt64([regex]::Match($node.InstanceId, 'DEV_([0-9A-F]{12})').Groups[1].Value, 16)
}

function Close-Joro {
    if ($script:Joro) {
        foreach ($obj in @($script:Joro.Service, $script:Joro.Device)) {
            # Windows PowerShell can't always Dispose WinRT objects; release is best-effort.
            if ($obj) { try { $obj.Dispose() } catch { } }
        }
    }
    $script:Joro = $null
}

function Connect-Joro {
    Close-Joro
    $bt = 'Windows.Devices.Bluetooth'
    $gatt = "$bt.GenericAttributeProfile"

    $device = Wait-JoroAsync ([Windows.Devices.Bluetooth.BluetoothLEDevice]::FromBluetoothAddressAsync((Get-JoroAddress))) ([type]"$bt.BluetoothLEDevice")
    if (-not $device) { throw 'Could not open the Joro over Bluetooth.' }
    $service = $null
    try {
        $services = Wait-JoroAsync ($device.GetGattServicesForUuidAsync($script:JoroServiceUuid)) ([type]"$gatt.GattDeviceServicesResult")
        Assert-JoroStatus $services.Status
        if ($services.Services.Count -eq 0) { throw 'Lighting service not found on this keyboard.' }
        $service = $services.Services[0]

        $open = Wait-JoroAsync ($service.OpenAsync([Windows.Devices.Bluetooth.GenericAttributeProfile.GattSharingMode]::SharedReadAndWrite)) ([type]"$gatt.GattOpenStatus")
        if ([string]$open -in @('SharingViolation', 'AccessDenied')) { throw $script:JoroLockMessage }
        if ([string]$open -notin @('Success', 'AlreadyOpened')) { throw "Could not open the lighting service ($open)." }

        $chars = Wait-JoroAsync ($service.GetCharacteristicsAsync()) ([type]"$gatt.GattCharacteristicsResult")
        Assert-JoroStatus $chars.Status
        $write  = $chars.Characteristics | Where-Object { $_.Uuid -eq $script:JoroWriteUuid } | Select-Object -First 1
        $notify = $chars.Characteristics | Where-Object { $_.Uuid -eq $script:JoroNotifyUuid } | Select-Object -First 1
        if (-not $write) { throw 'Lighting command channel not found.' }

        # Synapse subscribes to replies before sending commands; mirror that.
        if ($notify) {
            $null = Wait-JoroAsync ($notify.WriteClientCharacteristicConfigurationDescriptorAsync(
                [Windows.Devices.Bluetooth.GenericAttributeProfile.GattClientCharacteristicConfigurationDescriptorValue]::Notify)) ([type]"$gatt.GattCommunicationStatus")
        }
        $script:Joro = @{ Device = $device; Service = $service; Write = $write }
    }
    catch {
        foreach ($obj in @($service, $device)) { if ($obj) { try { $obj.Dispose() } catch { } } }
        throw
    }
}

function Write-JoroPacket([byte[]]$Bytes) {
    $buffer = [System.Runtime.InteropServices.WindowsRuntime.WindowsRuntimeBufferExtensions]::AsBuffer($Bytes)
    $result = Wait-JoroAsync ($script:Joro.Write.WriteValueWithResultAsync($buffer, [Windows.Devices.Bluetooth.GenericAttributeProfile.GattWriteOption]::WriteWithResponse)) ([Windows.Devices.Bluetooth.GenericAttributeProfile.GattWriteResult])
    Assert-JoroStatus $result.Status
}

function Send-JoroCommand([byte]$Class, [byte]$Command, [byte[]]$Data) {
    $script:JoroTxId = ($script:JoroTxId + 1) -band 0xFF
    Write-JoroPacket ([byte[]]@($script:JoroTxId, $Data.Length, 0, 0, $Class, $Command, 1, 0))
    Write-JoroPacket $Data
    Start-Sleep -Milliseconds 50
}

function Invoke-JoroLighting([string]$Color, [int]$Brightness) {
    if (-not $script:Joro) { Connect-Joro }
    if ($Color) {
        $hex = $Color.TrimStart('#')
        $r = [Convert]::ToByte($hex.Substring(0, 2), 16)
        $g = [Convert]::ToByte($hex.Substring(2, 2), 16)
        $b = [Convert]::ToByte($hex.Substring(4, 2), 16)
        Send-JoroCommand 0x10 0x03 ([byte[]]@(1, 0, 0, 1, $r, $g, $b))
    }
    if ($Brightness -ge 0) {
        Send-JoroCommand 0x10 0x05 ([byte[]]@([byte][Math]::Round([Math]::Min(100, $Brightness) * 255 / 100)))
    }
}

<#
.SYNOPSIS
    Apply a colour (hex RRGGBB) and/or brightness (0-100, -1 to skip). Reuses the open connection
    and reconnects once if it went stale (e.g. keyboard slept).
#>
function Set-JoroLighting([string]$Color, [int]$Brightness = -1) {
    try {
        Invoke-JoroLighting $Color $Brightness
    }
    catch {
        Close-Joro
        if ($_.Exception.Message -eq $script:JoroLockMessage) { throw }
        Invoke-JoroLighting $Color $Brightness
    }
}
