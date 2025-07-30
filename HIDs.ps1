Add-Type -AssemblyName System.Windows.Forms

# Create Form
$formHIDLookup = New-Object System.Windows.Forms.Form
$formHIDLookup.Text = "HID Sorting"
$formHIDLookup.Size = New-Object System.Drawing.Size(600, 500)
$formHIDLookup.StartPosition = "CenterScreen"

# Devices label
$labelDevices = New-Object System.Windows.Forms.Label
$labelDevices.Text = "Detected active devices:"
$labelDevices.Location = '10,10'
$labelDevices.Size = '550,20'
$formHIDLookup.Controls.Add($labelDevices)

# Devices listbox
$listDevices = New-Object System.Windows.Forms.ListBox
$listDevices.Location = '10,35'
#$listDevices.Size = '550,100'
$listDevices.Anchor = 'Top, Left, Right'
$listDevices.width = $formHIDLookup.Size.Width - 150
$listDevices.Height = $formHIDLookup.Size.Height - 180
$listDevices.HorizontalScrollbar = $true
$formHIDLookup.Controls.Add($listDevices)

# Up button
$buttonUp = New-Object System.Windows.Forms.Button
$buttonUp.Text = "Up"
#$buttonUp.Location = '220,175'
$buttonUp.Top = 35
$buttonUp.Left = $listDevices.Right + 10
$buttonUp.Anchor = 'Top, Right'
$buttonUp.Size = '60,30'
$formHIDLookup.Controls.Add($buttonUp)

# Down button
$buttonDown = New-Object System.Windows.Forms.Button
$buttonDown.Text = "Down"
#$buttonDown.Location = '220,215'
$buttonDown.Top = 75
$buttonDown.Left = $listDevices.Right + 10
$buttonDown.Anchor = 'Top, Right'
$buttonDown.Size = '60,30'
$formHIDLookup.Controls.Add($buttonDown)

# Status label
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Text = ""
$labelStatus.Location = '10,210'
$labelStatus.Size = '550,40'
$labelStatus.ForeColor = 'Red'
$formHIDLookup.Controls.Add($labelStatus)

# Action button
$buttonAction = New-Object System.Windows.Forms.Button
$buttonAction.Text = "Apply"
#$buttonAction.Location = '10,260'
$buttonAction.Size = '100,30'
$buttonAction.Top = $formHIDLookup.Size.Height - 150
$buttonAction.Left = 10
$buttonAction.Anchor = 'Bottom, Left'
$formHIDLookup.Controls.Add($buttonAction)

$buttonRefresh = New-Object System.Windows.Forms.Button
$buttonRefresh.Text = "Refresh Devices"
$buttonRefresh.Location = '120,260'
$buttonRefresh.Anchor = 'Bottom, Left'
$buttonRefresh.Top = $formHIDLookup.Size.Height - 150
$buttonRefresh.Left = $buttonAction.Right + 110
$buttonRefresh.Size = '150,30'
$formHIDLookup.Controls.Add($buttonRefresh)


# Global variables
#$devices = @()
$script:deviceList = @()

function LoadDevices {
    $oemName = ""
    $VID = ""
    $PID = ""
    $HID = ""
    $listDevices.Items.Clear()
    $devices = Get-PnpDevice -Class "HIDClass" | Where-Object {
        #$_.FriendlyName -like "*HID-compliant game controller*" -and $_.Status -eq "OK" -and $_.InstanceId -notlike "*HIDCLASS*"   #filters out vjoy or other virtual devices
        $_.FriendlyName -like "*HID-compliant game controller*" -and $_.Status -eq "OK"
    }
    Write-Host "Device count: $($devices.Count)"
    if ($devices.Count -eq 0) {
        $labelStatus.Text = "No active HID-compliant game controllers found."
        $buttonAction.Enabled = $false
        $buttonRefresh.Enabled = $true
    } else {
        $i = 1
        foreach ($d in $devices) {
            $instanceIdShort = $d.InstanceId
            if ($instanceIdShort -like "HID\*") {
                $instanceIdShort = $instanceIdShort.Substring(4)
                if ($instanceIdShort.Contains("\")) {
                    $instanceIdShort = $instanceIdShort.Split('\')[0]
                    # Extract VID and PID using regex
                    if ($instanceIdShort -match "VID_([0-9A-Fa-f]{4})&PID_([0-9A-Fa-f]{4})") {
                        $HIDVID = $matches[1]
                        $HIDPID = $matches[2]
                    } else {
                        $HIDVID = ""
                        $HIDPID = ""
                    }
                    $HID = "$HIDVID $HIDPID"

                }
            }
            $oemRegPath = "HKCU:\System\CurrentControlSet\Control\MediaProperties\PrivateProperties\Joystick\OEM\$($instanceIdShort)"
            #Write-Host "Checking OEM registry path: $oemRegPath"
            $oemName = "Unknown"
            if (Test-Path $oemRegPath) {
                $oemName = (Get-ItemProperty -Path $oemRegPath -Name OEMName ).OEMName
            } else {
                if ($i -eq 1) {
                    $oemName = "vJoy"
                } elseif ($i -eq 2) {
                    $oemName = "Gamepad"
                }
                Write-Host "OEM registry path not found for device: $($d.InstanceId) .This might be a vJoy or virtual device."
            }
            if ($HID -eq "") {
                $HID = "Unknown"
            }
            $listDevices.Items.Add("$i. $oemName - $HID [$($d.InstanceId)]")
            $i++
        }
        $buttonAction.Enabled = $true
        $buttonRefresh.Enabled = $true
    }
    $script:deviceList = $devices
    
}



$formHIDLookup.Add_Shown({
    $devices = LoadDevices
    $labelDevices.Text = "Detected $($devices.Count) devices."
})
$buttonAction.Add_Click({
    # Always reload devices to ensure count is up to date
    $devices = LoadDevices
    $orderInput = $textOrder.Text
    $order = $orderInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" } | ForEach-Object { $_ -as [int] }
    # Move vJoy (device without registry path) to the front of the array order
    $vjoyIndex = -1
    for ($i = 0; $i -lt $devices.Count; $i++) {
        $instanceIdShort = $devices[$i].InstanceId
        if ($instanceIdShort -like "HID\*") {
            $instanceIdShort = $instanceIdShort.Substring(4)
            if ($instanceIdShort.Contains("\")) {
                $instanceIdShort = $instanceIdShort.Split('\')[0]
            }
        }
        $oemRegPath = "HKCU:\System\CurrentControlSet\Control\MediaProperties\PrivateProperties\Joystick\OEM\$($instanceIdShort)"
        if (-not (Test-Path $oemRegPath)) {
            $vjoyIndex = $i + 1 # +1 because order is 1-based
            break
        }
    }
    if ($vjoyIndex -gt 0) {
        $order = @($vjoyIndex) + ($order | Where-Object { $_ -ne $vjoyIndex })
    }
    if ($order.Count -ne $devices.Count -or $order -contains $null -or ($order | Sort-Object | Get-Unique).Count -ne $devices.Count -or ($order | Where-Object { $_ -lt 1 -or $_ -gt $devices.Count }).Count -gt 0) {
        $labelStatus.Text = "Invalid order entered. Detected $($devices.Count) devices, but got $($order.Count) entries. Try again."
        return
    }
    # Disable all devices
    $labelStatus.ForeColor = 'Red'
    $labelStatus.Text = "Disabling all devices..."
    # Skip disabling device at position 1, only disable devices from position 2 onwards
    for ($i = 1; $i -lt $devices.Count; $i++) {
        $device = $devices[$i]
        if ($null -eq $device.InstanceId) {
            continue
        }
        $instanceIdShort = $device.InstanceId
        if ($instanceIdShort -like "HID\*") {
            $instanceIdShort = $instanceIdShort.Substring(4)
            if ($instanceIdShort.Contains("\")) {
                $instanceIdShort = $instanceIdShort.Split('\')[0]
            }
        }
        $oemRegPath = "HKCU:\System\CurrentControlSet\Control\MediaProperties\PrivateProperties\Joystick\OEM\$($instanceIdShort)"
        if (Test-Path $oemRegPath) {
            Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false
        }
    }
    Start-Sleep -Seconds 2
    # Enable in order
    $labelStatus.Text = "Enabling devices in the specified order..."
    foreach ($idx in $order) {
        $selectedDevice = $devices[$idx - 1]
        if ($null -eq $selectedDevice -or $null -eq $selectedDevice.InstanceId) {
            $labelStatus.ForeColor = 'Red'
            $labelStatus.Text = "Error: Device at position $idx is not valid or missing."
            return
        }
        Enable-PnpDevice -InstanceId $selectedDevice.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
    }
    $labelStatus.ForeColor = 'Green'
})

$buttonRefresh.Add_Click({
    $devices = LoadDevices
    $labelStatus.Text = "Devices refreshed. Detected $($devices.Count) devices."
})

[void]$formHIDLookup.ShowDialog()
