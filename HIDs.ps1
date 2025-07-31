Add-Type -AssemblyName System.Windows.Forms

$script:ScaleMultiplier = 1.0
<#       We'll use the screen dimensions below for suggesting a max window size                   #>
function Get-MaxScreenResolution {
    Add-Type -AssemblyName System.Windows.Forms
    $screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
    $screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height
    return "$screenWidth x $screenHeight"
}
#if ($debug) {Get-MaxScreenResolution}
function Get-DesktopResolutionScale {
    Add-Type -AssemblyName System.Windows.Forms
    $graphics = [System.Drawing.Graphics]::FromHwnd([System.IntPtr]::Zero)
    $desktopDpiX = $graphics.DpiX
    $scaleFactor = $desktopDpiX / 96  # 96 DPI is the default scale (100%)
    switch ($scaleFactor) {
        1 { $script:ScaleMultiplier = 1.0; return "100%" }
        1.25 { $script:ScaleMultiplier = 1.25; return "125%" }
        1.5 { $script:ScaleMultiplier = 1.5; return "150%" }
        1.75 { $script:ScaleMultiplier = 1.75; return "175%" }
        2 { $script:ScaleMultiplier = 2.0; return "200%" }
        default { $script:ScaleMultiplier = [math]::Round($scaleFactor * 100) / 100; return "$([math]::Round($scaleFactor * 100))%" }
    }
}Get-DesktopResolutionScale | Out-Null
if ($debug) {
    write-host "Resolution Scale: " (Get-DesktopResolutionScale)
    Write-Host "Scale Multiplier: " $script:ScaleMultiplier -BackgroundColor White -ForegroundColor Black
    Write-Host "Max Screen Resolution: " (Get-MaxScreenResolution) -BackgroundColor White -ForegroundColor Black
}

# Create Form
$formHIDLookup = New-Object System.Windows.Forms.Form
$formHIDLookup.Text = "HID Sorting"
$formHIDLookup.Size = New-Object System.Drawing.Size(600, 500)
$formHIDLookup.StartPosition = "CenterScreen"

# Devices label
$labelDevices = New-Object System.Windows.Forms.Label
$labelDevices.Text = "Detected active devices:"
$labelDevices.Location = New-Object System.Drawing.Point(10,10)
$labelDevices.Size = New-Object System.Drawing.Size(550,20)
$formHIDLookup.Controls.Add($labelDevices)

# Devices listbox
$listDevices = New-Object System.Windows.Forms.ListBox
$listDevices.Location = New-Object System.Drawing.Point(10,35)
#$listDevices.Size = New-Object System.Drawing.Size(550,100)
$listDevices.Anchor = 'Top, Left, Right'
$listDevices.Width = $formHIDLookup.Size.Width - 150
$listDevices.Height = $formHIDLookup.Size.Height - 180
$listDevices.HorizontalScrollbar = $true
$formHIDLookup.Controls.Add($listDevices)

# Up button
$buttonUp = New-Object System.Windows.Forms.Button
$buttonUp.Text = "Up"
#$buttonUp.Location = '220,175'
$buttonUp.Top = 35
$buttonUp.Left = $listDevices.Location.X + $listDevices.Width + 10
$buttonUp.Anchor = 'Top, Right'
$buttonUp.Size = New-Object System.Drawing.Size(60,30)
$formHIDLookup.Controls.Add($buttonUp)

# Down button
$buttonDown = New-Object System.Windows.Forms.Button
$buttonDown.Text = "Down"
#$buttonDown.Location = '220,215'
$buttonDown.Top = 75
$buttonDown.Left = $listDevices.Location.X + $listDevices.Width + 10
$buttonDown.Anchor = 'Top, Right'
$buttonDown.Size = New-Object System.Drawing.Size(60,30)
$formHIDLookup.Controls.Add($buttonDown)

$statusBar = New-Object System.Windows.Forms.StatusBar
$statusBar.Text = "Ready"
$statusBar.Dock = [System.Windows.Forms.DockStyle]::Bottom
$statusBar.Height = (20 * $script:ScaleMultiplier)
$statusBar.Font = New-Object System.Drawing.Font($statusBar.Font.FontFamily, [math]::Round($statusBar.Font.Size * $script:ScaleMultiplier), [System.Drawing.FontStyle]::Regular)
$statusBar.Name = "StatusBar"
$formHIDLookup.Controls.Add($statusBar)

# Action button
$buttonAction = New-Object System.Windows.Forms.Button
$buttonAction.Text = "Apply"
#$buttonAction.Location = '10,260'
$buttonAction.Size = New-Object System.Drawing.Size(100,30)
$buttonAction.Top = $formHIDLookup.Size.Height - 150
$buttonAction.Left = 10
$buttonAction.Anchor = 'Bottom, Left'
$formHIDLookup.Controls.Add($buttonAction)

$buttonRefresh = New-Object System.Windows.Forms.Button
$buttonRefresh.Text = "Refresh Devices"
$buttonRefresh.Location = New-Object System.Drawing.Point(120,260)
$buttonRefresh.Anchor = 'Bottom, Left'
$buttonRefresh.Top = $formHIDLookup.Size.Height - 150
$buttonRefresh.Left = $buttonAction.Left + $buttonAction.Width + 110
$buttonRefresh.Size = New-Object System.Drawing.Size(150,30)
$formHIDLookup.Controls.Add($buttonRefresh)


# Global variables
#$devices = @()
$script:deviceList = @()

function LoadDevices {
    $oemName = ""
    #$HID = ""
    $listDevices.Items.Clear()
    $devices = Get-PnpDevice -Class "HIDClass" | Where-Object {
        #$_.FriendlyName -like "*HID-compliant vendor-defined device*" -and $_.Status -eq "OK" # other HID devices
        $_.FriendlyName -like "*HID-compliant game controller*" -and $_.Status -eq "OK" -and $_.InstanceId -notlike "*HIDCLASS*"   #filters out vjoy or other virtual devices
        #$_.FriendlyName -like "*HID-compliant game controller*" -and $_.Status -eq "OK"
    }
    Write-Host "Device count: $($devices.Count)"
    if ($devices.Count -eq 0) {
        $statusBar.Text = "No active HID-compliant game controllers found."
        $buttonAction.Enabled = $false
        $buttonRefresh.Enabled = $true
    } else {
        $i = 1
        foreach ($d in $devices) {
            $instanceIdShort = $d.InstanceId
            $HIDVID = ""
            $HIDPID = ""
            $HID = ""
            if ($instanceIdShort -like "HID\*") {
                $instanceIdShort = $instanceIdShort.Substring(4)
                if ($instanceIdShort.Contains("\")) {
                    $instanceIdShort = $instanceIdShort.Split('\')[0]
                    # Extract VID and PID using regex
                    if ($instanceIdShort -match "VID_([0-9A-Fa-f]{4})&PID_([0-9A-Fa-f]{4})") {
                        $HIDVID = $matches[1]
                        $HIDPID = $matches[2]
                    }
                    $HID = "VendorID:$HIDVID ProductID:$HIDPID"
                }
            }
            $oemRegPath = "HKCU:\System\CurrentControlSet\Control\MediaProperties\PrivateProperties\Joystick\OEM\$($instanceIdShort)"
            $oemName = "Unknown"    #placeholder for OEM name
            if (Test-Path $oemRegPath) {
                $oemName = (Get-ItemProperty -Path $oemRegPath -Name OEMName ).OEMName
            } else {
                if ($HID -eq "") {
                    $HID = "Unknown"
                }
                Write-Host "OEM registry path not found for device: $($d.InstanceId) .This might be a vJoy or virtual device."
            }
            $listDevices.Items.Add("$i. $oemName - $HID [$($d.InstanceId)]")
            $i++
        }
        $buttonAction.Enabled = $true
        $buttonRefresh.Enabled = $true
    }
    $script:deviceList = $devices
    return $devices
}

# Helper: Get current device order from the ListBox
function Get-DeviceOrderFromListBox {
    $order = @()
    foreach ($item in $listDevices.Items) {
        # Each item is like "1. OEMName - HID [InstanceId]"
        if ($item -match "\[(.+?)\]$") {
            $instanceId = $matches[1]
            $instanceIdArray = $script:deviceList | ForEach-Object { $_.InstanceId }
            $idx = $instanceIdArray.IndexOf($instanceId)
            if ($idx -ge 0) {
                $order += ($idx + 1) # 1-based index
            }
        }
    }
    return $order
}

$buttonUp.Add_Click({
    $selectedIndex = $listDevices.SelectedIndex
    if ($selectedIndex -gt 0) {
        $temp = $listDevices.Items[$selectedIndex - 1]
        $listDevices.Items[$selectedIndex - 1] = $listDevices.Items[$selectedIndex]
        $listDevices.Items[$selectedIndex] = $temp
        $listDevices.SelectedIndex = $selectedIndex - 1
    }
})

$buttonDown.Add_Click({
    $selectedIndex = $listDevices.SelectedIndex
    if ($selectedIndex -lt $listDevices.Items.Count - 1) {
        $temp = $listDevices.Items[$selectedIndex + 1]
        $listDevices.Items[$selectedIndex + 1] = $listDevices.Items[$selectedIndex]
        $listDevices.Items[$selectedIndex] = $temp
        $listDevices.SelectedIndex = $selectedIndex + 1
    }
})

$formHIDLookup.Add_Shown({
    $devices = $null
    $devices = LoadDevices
    $labelDevices.Text = "Detected $($devices.Count) devices."
})

$buttonAction.Add_Click({
    # Always reload devices to ensure count is up to date
    $devices = $null
    $devices = LoadDevices
    Write-Host "Devices: $devices"
    #$orderInput = $textOrder.Text
    $order = Get-DeviceOrderFromListBox
    Write-Host "Order: $order"
    if ($order.Count -ne $devices.Count -or $order -contains $null -or ($order | Sort-Object | Get-Unique).Count -ne $devices.Count -or ($order | Where-Object { $_ -lt 1 -or $_ -gt $devices.Count }).Count -gt 0) {
        $statusBar.Text = "Invalid order entered. Detected $($devices.Count) devices, but got ordered list only has $($order.Count) entries. Try again."
        return
    }
    # Disable all devices
    $statusBar.ForeColor = 'Red'
    $statusBar.Text = "Disabling all devices..."
    # Disable all devices
    for ($i = 0; $i -lt $devices.Count; $i++) {
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
    $statusBar.Text = "Enabling devices in the specified order..."
    
    foreach ($idx in $order) {
        if ($idx -lt 1 -or $idx -gt $devices.Count) {
            $statusBar.ForeColor = 'Red'
            $statusBar.Text = "Error: Device position $idx is out of range."
            return
        }
        $selectedDevice = $devices[$idx - 1]
        if ($null -eq $selectedDevice -or $null -eq $selectedDevice.InstanceId) {
            $statusBar.ForeColor = 'Red'
            $statusBar.Text = "Error: Device at position $idx is not valid or missing."
            return
        }
        Enable-PnpDevice -InstanceId $selectedDevice.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
    }

    $statusBar.ForeColor = 'Green'
})

$buttonRefresh.Add_Click({
    $devices = $null
    $devices = LoadDevices
    $statusBar.Text = "Devices refreshed. Detected $($devices.Count) devices."
})

[void]$formHIDLookup.ShowDialog()
