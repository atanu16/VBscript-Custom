# ----------------------------
# Configuration
# ----------------------------
$siteUrl = "https://yourtenant.sharepoint.com/sites/Dev16"
$listName = "Temp"

# Login credentials
$user = "your_email@domain.com"
$plainPassword = "your_password"
$securePassword = ConvertTo-SecureString $plainPassword -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($user, $securePassword)

# Input data
$botName = "Credit Note Creation"
$platform = "BluePrism"
$username = "8TINMO.AA"
$machine = "MININT-R3007IK"
$status = "Completed"  # or "InProgress"
$region = "BNA"
$startTime = "2:00 PM"
$endTime = "4:00 PM"
$today = Get-Date -Format "yyyy-MM-dd"

# ----------------------------
# Connect to SharePoint
# ----------------------------
Connect-PnPOnline -Url $siteUrl -Credentials $cred

# ----------------------------
# Fetch existing item by Bot Name
# ----------------------------
$item = Get-PnPListItem -List $listName -PageSize 1000 | Where-Object { $_["Title"] -eq $botName }

if ($item) {
    if ($status -eq "Completed") {
        # Create new item
        Add-PnPListItem -List $listName -Values @{
            "Title"     = $botName
            "Platform"  = $platform
            "Username"  = $username
            "Machine"   = $machine
            "Date"      = $today
            "StartTime" = $startTime
            "EndTime"   = $endTime
            "Status"    = $status
            "Region"    = $region
        }
        Write-Host "New item created for COMPLETED bot."
    } else {
        # Update existing item
        Set-PnPListItem -List $listName -Identity $item.Id -Values @{
            "Platform"  = $platform
            "Username"  = $username
            "Machine"   = $machine
            "StartTime" = $startTime
            "EndTime"   = $endTime
            "Status"    = $status
            "Region"    = $region
        }
        Write-Host "Existing item updated for INPROGRESS bot."
    }
}
else {
    # Create new item if no match found
    Add-PnPListItem -List $listName -Values @{
        "Title"     = $botName
        "Platform"  = $platform
        "Username"  = $username
        "Machine"   = $machine
        "Date"      = $today
        "StartTime" = $startTime
        "EndTime"   = $endTime
        "Status"    = $status
        "Region"    = $region
    }
    Write-Host "New item created since no match found."
}