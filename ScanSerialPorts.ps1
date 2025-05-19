# Set output encoding to UTF-8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Query all serial ports using WMI
$ports = Get-WmiObject -Query "SELECT DeviceID, Caption FROM Win32_SerialPort"

# Check if any serial ports are found
if ($ports.Count -eq 0) {
    Write-Host "No available serial ports."
} else {
    # Display all serial port information
    foreach ($port in $ports) {
        Write-Host "Name: $($port.DeviceID), Description: $($port.Caption)"
    }

    # Ask the user if they want to remove occupied devices
    $response = Read-Host "Do you want to remove occupied devices? (Type 'Yes' to confirm, any other key to skip)"
    if ($response -eq "Yes") {
        foreach ($port in $ports) {
            try {
                Write-Host "Attempting to remove device $($port.DeviceID)..."
                $query = "SELECT * FROM Win32_PnPEntity WHERE DeviceID = '$($port.DeviceID.Replace('\\', '\\\\'))'"
                $device = Get-WmiObject -Query $query
                if ($device) {
                    $device.Remove()  # Attempt to remove the device
                    Write-Host "Device removed: $($port.Caption)"
                } else {
                    Write-Host "Unable to find device: $($port.DeviceID)"
                }
            } catch {
                Write-Host "Failed to remove device: $($_.Exception.Message)"
            }
        }
    } else {
        Write-Host "No removal operations were performed."
    }
}

# Display all available COM port names
Write-Host "`nAvailable Port Names:"
$serialPorts = [System.IO.Ports.SerialPort]::GetPortNames()
if ($serialPorts.Count -eq 0) {
    Write-Host "No available serial ports."
} else {
    $serialPorts | ForEach-Object { Write-Host $_ }
}


# Virtual COM Port information
Write-Host "If the above is 'No available serial ports.' but below shows available COM ports, these are virtual COM ports."


# Keep the window open and prompt user to press any key to exit
Write-Host "`nOperation completed, press any key to exit..."
Read-Host
