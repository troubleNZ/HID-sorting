Add-Type -AssemblyName System.Windows.Forms
$debug = $false
$script:ScaleMultiplier = 1.0

$checkIfAdmin = {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
if (-not ($checkIfAdmin.Invoke())) {
    Write-Host "[Read Only Mode]: Please Relaunch as Administrator if you wish to make changes." -ForegroundColor Red
}
function Get-MaxScreenResolution {
    $screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
    $screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height
    return "$screenWidth x $screenHeight"
}
#if ($debug) {Get-MaxScreenResolution}
function Get-DesktopResolutionScale {
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
}

#Get-DesktopResolutionScale | Out-Null

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
$labelDevices.Font = New-Object System.Drawing.Font($labelDevices.Font.FontFamily, [math]::Round($labelDevices.Font.Size * $script:ScaleMultiplier), [System.Drawing.FontStyle]::Regular)
$formHIDLookup.Controls.Add($labelDevices)

# Devices listbox
$listDevices = New-Object System.Windows.Forms.ListBox
$listDevices.Location = New-Object System.Drawing.Point(10,35)
$listDevices.Anchor = 'Top, Left, Right, Bottom'
$listDevices.Width = $formHIDLookup.Size.Width - 150
$listDevices.Height = $formHIDLookup.Size.Height - 180
$listDevices.HorizontalScrollbar = $true
$listDevices.Font = New-Object System.Drawing.Font($listDevices.Font.FontFamily, [math]::Round($listDevices.Font.Size * $script:ScaleMultiplier), [System.Drawing.FontStyle]::Regular)
$formHIDLookup.Controls.Add($listDevices)

# Make $listDevices resize with the form
$formHIDLookup.Add_Resize({
    $listDevices.Width = $formHIDLookup.ClientSize.Width - 150
    $listDevices.Height = $formHIDLookup.ClientSize.Height - 180
})

# Up button
$buttonUp = New-Object System.Windows.Forms.Button
$buttonUp.Text = "Up"
#$buttonUp.Location = '220,175'
$buttonUp.Top = 35
$buttonUp.Left = $listDevices.Location.X + $listDevices.Width + 10
$buttonUp.Anchor = 'Top, Right'
$buttonUp.Size = New-Object System.Drawing.Size(60,30)
$buttonUp.Enabled = $false
$buttonUp.Font = New-Object System.Drawing.Font($buttonUp.Font.FontFamily, [math]::Round($buttonUp.Font.Size * $script:ScaleMultiplier), [System.Drawing.FontStyle]::Regular)
$formHIDLookup.Controls.Add($buttonUp)

# Down button
$buttonDown = New-Object System.Windows.Forms.Button
$buttonDown.Text = "Down"
#$buttonDown.Location = '220,215'
$buttonDown.Top = 75
$buttonDown.Left = $listDevices.Location.X + $listDevices.Width + 10
$buttonDown.Anchor = 'Top, Right'
$buttonDown.Size = New-Object System.Drawing.Size(60,30)
$buttonDown.Enabled = $false
$buttonDown.Font = New-Object System.Drawing.Font($buttonDown.Font.FontFamily, [math]::Round($buttonDown.Font.Size * $script:ScaleMultiplier), [System.Drawing.FontStyle]::Regular)
$formHIDLookup.Controls.Add($buttonDown)

$listDevices.Add_SelectedIndexChanged({
    $selectedIndex = $listDevices.SelectedIndex
    $count = $listDevices.Items.Count

    # Enable Up if not the first item and something is selected
    $buttonUp.Enabled = ($selectedIndex -gt 0)
    # Enable Down if not the last item and something is selected
    $buttonDown.Enabled = ($selectedIndex -ge 0 -and $selectedIndex -lt ($count - 1))
})

$statusBar = New-Object System.Windows.Forms.StatusBar
$statusBar.Text = "Ready"
$statusBar.Dock = [System.Windows.Forms.DockStyle]::Bottom
$statusBar.Height = (20 * $script:ScaleMultiplier)
$statusBar.Font = New-Object System.Drawing.Font($statusBar.Font.FontFamily, [math]::Round($statusBar.Font.Size * $script:ScaleMultiplier), [System.Drawing.FontStyle]::Regular)
$statusBar.Name = "StatusBar"
$formHIDLookup.Controls.Add($statusBar)

# Add tooltip for $listDevices
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($listDevices, "Right Click to copy Device HID information to clipboard.")

# MouseDown event for right/left click actions
$listDevices.Add_MouseDown({
    param($sender, $e)
    $index = $listDevices.IndexFromPoint($e.Location)
    if ($index -ge 0) {
        $itemText = $listDevices.Items[$index]
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
            [Windows.Forms.Clipboard]::SetText($itemText)
            $statusBar.Text = "Copied to clipboard: $itemText"
        }
        elseif ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            $statusBar.Text = $itemText
        }
    }
})



# Action button
$buttonAction = New-Object System.Windows.Forms.Button
$buttonAction.Text = "Apply"
#$buttonAction.Location = '10,260'
$buttonAction.Size = New-Object System.Drawing.Size(100,30)
$buttonAction.Top = $listDevices.Location.Y + $listDevices.Height + 10
$buttonAction.Left = 10
$buttonAction.Anchor = 'Bottom, Left'
$buttonAction.Font = New-Object System.Drawing.Font($buttonAction.Font.FontFamily, [math]::Round($buttonAction.Font.Size * $script:ScaleMultiplier), [System.Drawing.FontStyle]::Regular)
$formHIDLookup.Controls.Add($buttonAction)

$buttonRefresh = New-Object System.Windows.Forms.Button
$buttonRefresh.Text = "Refresh Devices"
$buttonRefresh.Location = New-Object System.Drawing.Point(120,260)
$buttonRefresh.Anchor = 'Bottom, Left'
$buttonRefresh.Top = $listDevices.Location.Y + $listDevices.Height + 10
$buttonRefresh.Left = $buttonAction.Left + $buttonAction.Width + 110
$buttonRefresh.Size = New-Object System.Drawing.Size(150,30)
$buttonRefresh.Font = New-Object System.Drawing.Font($buttonRefresh.Font.FontFamily, [math]::Round($buttonRefresh.Font.Size * $script:ScaleMultiplier), [System.Drawing.FontStyle]::Regular)
$formHIDLookup.Controls.Add($buttonRefresh)

$script:globalDeviceList = @()

$script:IncludeVirtualDevices = $false
$script:IncludeDisconnected = $false

# Checkbox to include virtual devices
# Radio buttons for device filter options
$radioActiveOnly = New-Object System.Windows.Forms.RadioButton
$radioActiveOnly.Text = "Active Only"
$radioActiveOnly.AutoSize = $true
$radioActiveOnly.Top = $buttonRefresh.Top - 5
$radioActiveOnly.Left = $buttonRefresh.Left + $buttonRefresh.Width + 20
$radioActiveOnly.Anchor = 'Bottom, Left'
$radioActiveOnly.Font = New-Object System.Drawing.Font($radioActiveOnly.Font.FontFamily, [math]::Round($radioActiveOnly.Font.Size * $script:ScaleMultiplier), [System.Drawing.FontStyle]::Regular)
$radioActiveOnly.Checked = $true

$radioIncludeVirtual = New-Object System.Windows.Forms.RadioButton
$radioIncludeVirtual.Text = "Include virtual devices"
$radioIncludeVirtual.AutoSize = $true
$radioIncludeVirtual.Top = $radioActiveOnly.Top + $radioActiveOnly.Height + 5
$radioIncludeVirtual.Left = $radioActiveOnly.Left
$radioIncludeVirtual.Anchor = 'Bottom, Left'
$radioIncludeVirtual.Font = New-Object System.Drawing.Font($radioIncludeVirtual.Font.FontFamily, [math]::Round($radioIncludeVirtual.Font.Size * $script:ScaleMultiplier), [System.Drawing.FontStyle]::Regular)

$radioIncludeDisconnected = New-Object System.Windows.Forms.RadioButton
$radioIncludeDisconnected.Text = "Include Disconnected"
$radioIncludeDisconnected.AutoSize = $true
$radioIncludeDisconnected.Top = $radioIncludeVirtual.Top + $radioIncludeVirtual.Height + 5
$radioIncludeDisconnected.Left = $radioActiveOnly.Left
$radioIncludeDisconnected.Anchor = 'Bottom, Left'
$radioIncludeDisconnected.Font = New-Object System.Drawing.Font($radioIncludeDisconnected.Font.FontFamily, [math]::Round($radioIncludeDisconnected.Font.Size * $script:ScaleMultiplier), [System.Drawing.FontStyle]::Regular)

# Add radio buttons to the form
$formHIDLookup.Controls.Add($radioActiveOnly)
$formHIDLookup.Controls.Add($radioIncludeVirtual)
$formHIDLookup.Controls.Add($radioIncludeDisconnected)

# Radio button logic
$radioActiveOnly.Add_CheckedChanged({
    if ($radioActiveOnly.Checked) {
        $script:IncludeVirtualDevices = $false
        $script:IncludeDisconnected = $false
    }
})
$radioIncludeVirtual.Add_CheckedChanged({
    if ($radioIncludeVirtual.Checked) {
        $script:IncludeVirtualDevices = $true
        $script:IncludeDisconnected = $false
    }
})
$radioIncludeDisconnected.Add_CheckedChanged({
    if ($radioIncludeDisconnected.Checked) {
        $script:IncludeVirtualDevices = $false
        $script:IncludeDisconnected = $true
    }
})

#$formHIDLookup.Controls.Add($checkboxVirtualDevices)

function LoadDevices {
    $oemName = ""
    #$HID = ""
    $listDevices.Items.Clear()
    $detectedDevices = Get-PnpDevice -Class "HIDClass" | Where-Object {
        if ($script:IncludeDisconnected) {
            return $_.FriendlyName -like "*HID-compliant game controller*"
        }
        if ($script:IncludeVirtualDevices) {
            return $_.FriendlyName -like "*HID-compliant game controller*" -and $_.Status -eq "OK"
        } else {
            return $_.FriendlyName -like "*HID-compliant game controller*" -and $_.Status -eq "OK" -and $_.InstanceId -notlike "*HIDCLASS*"    # filters out vjoy or other virtual devices
                #$_.FriendlyName -like "*HID-compliant game controller*" -and $_.Status -eq "OK"                                            # includes vJoy or other virtual devices
                #$_.FriendlyName -like "*HID-compliant game controller*"                                                                    # includes all HID-compliant game controllers
        }
    }
    if ($debug) { Write-Host "Device count: $($detectedDevices.Count)"}
    if ($detectedDevices.Count -eq 0) {
        $statusBar.Text = "No active HID-compliant game controllers found."
        $buttonAction.Enabled = $false
        $buttonRefresh.Enabled = $true
    } else {
        $i = 1
        foreach ($d in $detectedDevices) {
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
            } else {
                # If the InstanceId is not in the expected format, we can still use it
                $HID = $instanceIdShort
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
            # If HID is "Unknown", try to look up VID/PID in usb_ids.csv
            if ($oemName -eq "Unknown") {
                $basePath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
                $csvPath = Join-Path -Path $basePath -ChildPath "usb_ids.csv"
                if (Test-Path $csvPath) {
                    Write-Host "Attempting to look up VID/PID in usb_ids.csv : Vendor ID: $HIDVID Product ID: $HIDPID"
                    try {
                        $usbIds = Import-Csv -Path $csvPath
                        $match = $usbIds | Where-Object {
                            ($_.VendorID -eq $HIDVID -or $_.VendorID -eq ("0x" + $HIDVID)) -and
                            ($_.ProductID -eq $HIDPID -or $_.ProductID -eq ("0x" + $HIDPID))
                        }
                        if ($match) {
                            $oemName = $match.ProductName
                            write-host "Found OEM name in usb_ids.csv: $oemName"
                            #$HID = "VendorID:$($match.VendorID) ProductID:$($match.ProductID)"
                        }
                    } catch {
                        Write-Host "Error reading usb_ids.csv: $($_.Exception.Message)"
                    }
                } else {
                    Write-Host "usb_ids.csv not found.`nTo generate this file, please check the 'ConvertUSB_IDs_toCSV.ps1' script in this repository.`nIt contains a link to download the raw USB ID data and instructions to run the script to create usb_ids.csv.`nThis file is only used as a secondary source for hardware IDs, and you can carry on without it."
                }
            }

            $listDevices.Items.Add("$i. $oemName - $HID [$($d.InstanceId)]")
            $i++
        }
        $buttonAction.Enabled = $true
        $buttonRefresh.Enabled = $true
    }
    $script:globalDeviceList = $detectedDevices
    if ($debug) { Write-Host "globalDeviceList count: $($script:globalDeviceList.Count)"}
    if ($debug) { Write-Host "globalDeviceList contents: $($script:globalDeviceList | ForEach-Object { $_.InstanceId })"}
    return $detectedDevices

}

function Get-DeviceOrderFromListBox {
    $order = @()
    foreach ($item in $listDevices.Items) {
        # Each item is like "1. OEMName - HID [InstanceId]"
        if ($item -match "\[(.+?)\]$") {
            $instanceId = $matches[1]
            $instanceIdArray = $script:globalDeviceList | ForEach-Object { $_.InstanceId }
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
    LoadDevices | Out-Null
    #$devices = $script:globalDeviceList
    $labelDevices.Text = "Detected $($script:globalDeviceList.Count) active devices."
    if ($debug) { Write-Host "devices detected: $($script:globalDeviceList.Count)" }
    if ($script:globalDeviceList.Count -eq 0) {
        $statusBar.Text = "No active HID-compliant game controllers found."
        $buttonAction.Enabled = $false
        $buttonRefresh.Enabled = $true
    } else {
        $statusBar.Text = "Ready"
        $buttonAction.Enabled = $true
        $buttonRefresh.Enabled = $true
    }
    if (-not ($checkIfAdmin.Invoke())) {
    $statusBar.Text = "[Read Only Mode]: Please Relaunch as Administrator if you wish to make changes."
}
})

$buttonAction.Add_Click({
    # Always reload devices to ensure count is up to date
    #LoadDevices
    Write-Host "Devices: $script:globalDeviceList"
    $SortOrder = Get-DeviceOrderFromListBox
    Write-Host "Order: $SortOrder"
    if ($SortOrder.Count -ne $script:globalDeviceList.Count -or $SortOrder -contains $null -or ($SortOrder | Sort-Object | Get-Unique).Count -ne $script:globalDeviceList.Count -or ($SortOrder | Where-Object { $_ -lt 1 -or $_ -gt $script:globalDeviceList.Count }).Count -gt 0) {
        $statusBar.Text = "Invalid order entered. Detected $($script:globalDeviceList.Count) devices, but got ordered list only has $($SortOrder.Count) entries. Try again."
        return
    }
    # Disable all devices
    $statusBar.Text = "Disabling all devices..."
    # Disable all devices
    for ($i = 0; $i -lt $script:globalDeviceList.Count; $i++) {
        $device = $script:globalDeviceList[$i]
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
            try {
                Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction Continue
            } catch {
                $statusBar.Text = "Error disabling device $($device.InstanceId): $($_.Exception.Message)"
                Write-Host "Error disabling device $($device.InstanceId): $($_.Exception.Message)"
                return
            }
        }
    }
    Start-Sleep -Seconds 2
    # Enable in order
    $statusBar.Text = "Enabling devices in the specified order..."

    foreach ($idx in $SortOrder) {
        if ($idx -lt 1 -or $idx -gt $script:globalDeviceList.Count) {
            $statusBar.Text = "Error: Device position $idx is out of range."
            return
        }
        $selectedDevice = $script:globalDeviceList[$idx - 1]
        if ($null -eq $selectedDevice -or $null -eq $selectedDevice.InstanceId) {
            $statusBar.Text = "Error: Device at position $idx is not valid or missing."
            return
        }
        try {
            Enable-PnpDevice -InstanceId $selectedDevice.InstanceId -Confirm:$false -ErrorAction Continue
        } catch {
            $statusBar.Text = "Error enabling device $($selectedDevice.InstanceId): $($_.Exception.Message)"
            Write-Host "Error enabling device $($selectedDevice.InstanceId): $($_.Exception.Message)"
            return
        }
        Write-Host "Enabled device: $($selectedDevice.InstanceId) from position $idx"
        Start-Sleep -Seconds 1
    }
    $statusBar.Text = "Action completed successfully. Devices enabled in the specified order: $SortOrder"
})

$buttonRefresh.Add_Click({
    LoadDevices | Out-Null
    $statusBar.Text = "Devices refreshed. Detected $($script:globalDeviceList.Count) devices."
    $labelDevices.Text = "Detected $($script:globalDeviceList.Count) devices."
})

[void]$formHIDLookup.ShowDialog()
