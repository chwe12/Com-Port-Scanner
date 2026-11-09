<#
.SYNOPSIS
    List all COM ports and optionally remove phantom (unplugged) port devices.

.DESCRIPTION
    - Lists every device in the "Ports (COM & LPT)" class, including devices that
      are no longer connected, plus COM names that exist outside that class
      (for example some virtual COM port drivers).
    - Optionally checks whether each connected port is in use by another program.
    - Optionally removes phantom devices (Present = False) with pnputil, which frees
      the COM numbers they still hold. Connected devices are never removed.

.PARAMETER ListOnly
    Only print the list. No questions, nothing is opened or removed.

.NOTES
    Requires Windows PowerShell 5.1 or PowerShell 7 on Windows.
    Removing devices requires administrator rights and Windows 10 2004 (build 19041) or later.
    Keep this file ASCII-only: Windows PowerShell 5.1 reads BOM-less scripts with the ANSI code page.
#>
[CmdletBinding()]
param(
    [switch]$ListOnly
)

function Test-IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    return (New-Object Security.Principal.WindowsPrincipal $identity).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
}

# COM name of a device: registry "Device Parameters\PortName" first, then "(COMx)" in the friendly name
function Get-PortName {
    param($Device)
    $key = "HKLM:\SYSTEM\CurrentControlSet\Enum\$($Device.InstanceId)\Device Parameters"
    $name = (Get-ItemProperty -LiteralPath $key -Name PortName -ErrorAction SilentlyContinue).PortName
    if ($name) { return $name }
    if ($Device.FriendlyName -match '\((COM\d+|LPT\d+)\)') { return $Matches[1] }
    return ''
}

# Try to open the port: opened = Free, access denied = In use
function Test-PortInUse {
    param([string]$Name)
    $port = New-Object System.IO.Ports.SerialPort $Name
    try {
        $port.Open()
        $port.Close()
        return 'Free'
    } catch [System.UnauthorizedAccessException] {
        return 'In use'
    } catch {
        return 'Unknown'
    } finally {
        $port.Dispose()
    }
}

function Get-SerialPortList {
    $devices = @(Get-PnpDevice -Class Ports -ErrorAction SilentlyContinue)
    $rows = foreach ($d in $devices) {
        $portName = Get-PortName $d
        if ($portName -like 'LPT*') { continue }   # skip printer ports
        [PSCustomObject]@{
            Port       = $portName
            Present    = [bool]$d.Present
            Status     = if ($d.Present) { [string]$d.Status } else { 'Unplugged' }
            InUse      = '-'
            Name       = $d.FriendlyName
            InstanceId = $d.InstanceId
        }
    }
    $rows = @($rows)

    # COM names Windows reports that are not in the Ports class (some virtual COM drivers)
    $known = @($rows | ForEach-Object { $_.Port })
    foreach ($n in ([System.IO.Ports.SerialPort]::GetPortNames() | Sort-Object -Unique)) {
        if ($known -notcontains $n) {
            $rows += [PSCustomObject]@{
                Port       = $n
                Present    = $true
                Status     = 'Virtual/Other'
                InUse      = '-'
                Name       = '(not in "Ports (COM & LPT)", likely a virtual COM port)'
                InstanceId = ''
            }
        }
    }

    # Sort by COM number
    return @($rows | Sort-Object @{ Expression = {
        if ($_.Port -match '\d+') { [int]$Matches[0] } else { [int]::MaxValue } } })
}

function Show-PortTable {
    param($Rows)
    $Rows | Format-Table Port, Present, Status, InUse, Name -AutoSize | Out-String -Width 200 | Write-Host
}

function Remove-PhantomPorts {
    param($Rows)
    $phantoms = @($Rows | Where-Object { -not $_.Present -and $_.InstanceId })
    if ($phantoms.Count -eq 0) {
        Write-Host 'No unplugged (phantom) COM port devices found.'
        return
    }

    Write-Host "Unplugged (phantom) COM port devices: $($phantoms.Count)"
    $phantoms | Format-Table Port, Name, InstanceId -AutoSize | Out-String -Width 250 | Write-Host

    if (-not (Test-IsAdmin)) {
        Write-Host 'Removing devices requires administrator rights. Run RunSerialPorts.bat or an elevated PowerShell.' -ForegroundColor Yellow
        return
    }
    if ([Environment]::OSVersion.Version.Build -lt 19041) {
        Write-Host 'pnputil /remove-device needs Windows 10 2004 or later.' -ForegroundColor Yellow
        Write-Host 'Remove them manually: Device Manager > View > Show hidden devices > Ports (COM & LPT).'
        return
    }

    $answer = Read-Host 'Remove these unplugged devices? Connected devices are not touched. (Type Yes to confirm)'
    if ($answer -ne 'Yes') {
        Write-Host 'Skipped. Nothing was removed.'
        return
    }

    foreach ($p in $phantoms) {
        $output = & pnputil /remove-device "$($p.InstanceId)" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Removed: $($p.Port) $($p.Name)" -ForegroundColor Green
        } else {
            Write-Host "Failed:  $($p.Port) $($p.Name) (pnputil exit code $LASTEXITCODE)" -ForegroundColor Red
            $output | ForEach-Object { Write-Host "         $_" }
        }
    }
}

function Invoke-Main {
    Write-Host 'COM Port Scanner'
    Write-Host ''

    $rows = Get-SerialPortList
    if ($rows.Count -eq 0) {
        Write-Host 'No COM ports found.'
        return
    }

    if (-not $ListOnly) {
        Show-PortTable $rows
        Write-Host 'Check whether connected ports are in use? This briefly opens each free port,'
        Write-Host 'which can reset some boards (for example Arduino).'
        $check = Read-Host 'Type y to check, anything else to skip'
        if ($check -eq 'y') {
            foreach ($r in ($rows | Where-Object { $_.Present -and $_.Port -like 'COM*' })) {
                $r.InUse = Test-PortInUse $r.Port
            }
        }
    }

    Show-PortTable $rows
    if ($ListOnly) { return }

    Remove-PhantomPorts $rows
}

# Run only when executed directly (dot-sourcing loads the functions for testing)
if ($MyInvocation.InvocationName -ne '.') {
    Invoke-Main
    if (-not $ListOnly) {
        Write-Host ''
        Read-Host 'Done. Press Enter to exit' | Out-Null
    }
}
