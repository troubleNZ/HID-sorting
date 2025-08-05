# convert_usb_ids_localmerge.ps1
# Parses a locally provided usb.ids, merges with any existing csv (preserving user edits/additions),
# and emits usb_ids.csv plus audit of user-only extras.

# get latest version from site: http://www.linux-usb.org/usb.ids

# === Configuration ===
$IdsFile         = "usb.ids"
$CsvFile         = "usb_ids.csv"
$ExtraCsv        = "user_extra_entries.csv"    # entries only in existing CSV

# === Step 1: Ensure usb.ids exists ===
if (-not (Test-Path $IdsFile)) {
    Write-Error "Required file '$IdsFile' not found in current directory. Please download the latest usb.ids manually and place it here."
    exit 1
}

# === Step 2: Parse usb.ids into structured records ===
$vendors = @{}
$products = @{}
$interfaces = @()
$parsedRecords = @()

$curVendorID = $null
$curVendorName = $null
$curProductID = $null
$curProductName = $null

Get-Content $IdsFile | ForEach-Object {
    $line = $_
    if ($line -match '^\s*#') {
        return
    } elseif ($line -match '^([0-9A-Fa-f]{4})\s+(.+)') {
        # vendor
        $curVendorID = $matches[1]
        $curVendorName = $matches[2].Trim()
        $vendors[$curVendorID] = $curVendorName
        $curProductID = $null
        $curProductName = $null
    } elseif ($line -match '^\t([0-9A-Fa-f]{4})\s+(.+)') {
        # product
        $curProductID = $matches[1]
        $curProductName = $matches[2].Trim()
        $products["$curVendorID`:$curProductID"] = $curProductName

        $parsedRecords += [PSCustomObject]@{
            VendorID      = $curVendorID
            VendorName    = $curVendorName
            ProductID     = $curProductID
            ProductName   = $curProductName
            InterfaceID   = ""
            InterfaceName = ""
        }
    } elseif ($line -match '^\t\t([0-9A-Fa-f]{2})\s+(.+)') {
        # interface under current product
        if ($curVendorID -and $curProductID) {
            $ifaceID = $matches[1]
            $ifaceName = $matches[2].Trim()

            $interfaces += [PSCustomObject]@{
                VendorID      = $curVendorID
                VendorName    = $curVendorName
                ProductID     = $curProductID
                ProductName   = $curProductName
                InterfaceID   = $ifaceID
                InterfaceName = $ifaceName
            }

            $parsedRecords += [PSCustomObject]@{
                VendorID      = $curVendorID
                VendorName    = $curVendorName
                ProductID     = $curProductID
                ProductName   = $curProductName
                InterfaceID   = $ifaceID
                InterfaceName = $ifaceName
            }
        }
    }
}

# === Step 3: Load existing CSV (if any) and merge ===

function Get-Key($o) {
    $v = if ($o.PSObject.Properties.Match('VendorID')) { $o.VendorID } else { "" }
    $p = if ($o.PSObject.Properties.Match('ProductID')) { $o.ProductID } else { "" }
    $i = if ($o.PSObject.Properties.Match('InterfaceID')) { $o.InterfaceID } else { "" }
    return "$v|$p|$i"
}

$finalRecords = @()
$conflicts = @()
$userExtras = @()

# Build parsed map
$parsedMap = @{}
foreach ($r in $parsedRecords) {
    $key = Get-Key $r
    $parsedMap[$key] = $r
}

if (Test-Path $CsvFile) {
    Write-Output "Existing CSV detected. Loading and comparing..."

    try {
        $oldRecords = Import-Csv -Path $CsvFile
    } catch {
        Write-Warning "Failed to import existing CSV: $($_.Exception.Message)"
        $oldRecords = @()
    }

    # Normalize old records to ensure interface fields exist
    foreach ($o in $oldRecords) {
        if (-not $o.PSObject.Properties.Match('InterfaceID')) { $o | Add-Member -NotePropertyName InterfaceID -NotePropertyValue "" -Force }
        if (-not $o.PSObject.Properties.Match('InterfaceName')) { $o | Add-Member -NotePropertyName InterfaceName -NotePropertyValue "" -Force }
    }

    # Build old map
    $oldMap = @{}
    foreach ($o in $oldRecords) {
        $key = Get-Key $o
        $oldMap[$key] = $o
    }

    # Compare and merge
    foreach ($key in $oldMap.Keys) {
        if ($parsedMap.ContainsKey($key)) {
            $new = $parsedMap[$key]
            $old = $oldMap[$key]
            if (($new.VendorName -ne $old.VendorName) -or ($new.ProductName -ne $old.ProductName) -or ($new.InterfaceName -ne $old.InterfaceName)) {
                $conflicts += [PSCustomObject]@{
                    Key           = $key
                    ParsedVendor  = $new.VendorName
                    OldVendor     = $old.VendorName
                    ParsedProduct = $new.ProductName
                    OldProduct    = $old.ProductName
                    ParsedIface   = $new.InterfaceName
                    OldIface      = $old.InterfaceName
                }
                # Keep user’s existing version
                $finalRecords += $old
                $parsedMap.Remove($key) | Out-Null
            } else {
                $finalRecords += $new
                $parsedMap.Remove($key) | Out-Null
            }
        } else {
            # only in old CSV => user-added / divergent
            $userExtras += $oldMap[$key]
        }
    }

    # Add remaining authoritative parsed entries
    foreach ($remaining in $parsedMap.Values) {
        $finalRecords += $remaining
    }

    if ($userExtras.Count -gt 0) {
        $userExtras | Export-Csv -Path $ExtraCsv -NoTypeInformation -Encoding UTF8
        Write-Output "Preserving $($userExtras.Count) user-only entry(ies); archived to $ExtraCsv."
        foreach ($u in $userExtras) {
            $finalRecords += $u
        }
    }

} else {
    $finalRecords = $parsedRecords
}

# Deduplicate finalRecords
$seen = @{}
$deduped = @()
foreach ($r in $finalRecords) {
    $k = Get-Key $r
    if (-not $seen.ContainsKey($k)) {
        $deduped += $r
        $seen[$k] = $true
    }
}

# === Step 4: Export merged CSV ===
$deduped | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8

# === Step 5: Summary ===
$vendorCount = ($deduped | Select-Object -ExpandProperty VendorID | Sort-Object -Unique).Count
$productCount = ($deduped | ForEach-Object { "$($_.VendorID):$($_.ProductID)" } | Sort-Object -Unique).Count
$interfaceCount = ($deduped | Where-Object { $_.InterfaceID -ne "" } | Measure-Object).Count

Write-Output ""
Write-Output "FINAL SUMMARY:"
Write-Output "  Vendors:    $vendorCount"
Write-Output "  Products:   $productCount"
Write-Output "  Interfaces: $interfaceCount"
Write-Output "  Total Rows: $($deduped.Count)"

if ($conflicts.Count -gt 0) {
    Write-Output ""
    Write-Output "CONFLICTS (kept existing CSV names):"
    $conflicts | Format-Table -AutoSize
}

Write-Output ""
Write-Output "Merged CSV saved to: $CsvFile"
if (Test-Path $ExtraCsv) {
    Write-Output "User-only extras file: $ExtraCsv"
}
