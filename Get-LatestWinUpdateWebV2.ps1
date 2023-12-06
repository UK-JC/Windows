function Get-LatestWinUpdateWeb{
  param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("Win10", "Win11")]
    [string]$os,
    [Parameter(Mandatory=$true)]
    [ValidateSet("21h2", "22h2", "23h2")]
    [string]$build,
    [Parameter(Mandatory=$false)]
    [switch]$notes
  )

<#*************************************************************************************
Author: Jeffrey Carroll
### Requires Elevation to install the PowerHTML module if not already installed.

Function Usage:

-os : You can use Win10 or Win11 to request for specific operating systems
-build : You can use 22h2 or 21h2 for the builds for Windows 10 and Windows 11, but 22h3 is Windows 11 only.
-notes : $true will display the notes. Not using this parameter will set the value to
$false, which will display the update data only.

The following data will be return on an update:
 Servicing_Option    
 Availability_Date   
 Build              
 KB_Article        
 KB_Number 
Example 1 : Get-LatestWinUpdateWeb -os "Win10" -build "22H2"
This will return the windows update information

Example 2 : Get-LatestWinUpdateWeb -os "Win10" -build "22H2" -notes
This will return the windows update notes, showing what changes were made in the KB
*************************************************************************************#>


Install-Module PowerHTML

# Define the URL to scrape
if($os -eq "Win10"){
 $url = "https://learn.microsoft.com/en-us/windows/release-health/release-information"
} else{
 $url = "https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information"
}

$html = ConvertFrom-Html -URI $url

$items = $html.InnerHtml -split "<[^>]+>"

# Remove any empty entries
$items = $items | Where-Object { $_ -ne "" }

# Output the array of items
$string = $items -join ""
$array = $string -split "(?m)^\s*$"
$loop = 0
$ready = $false
$ready = $false
$latestUpdateItem = @()
$latestUpdate23h2 = $null
$latestUpdate22h2 = $null
$latestUpdate21h2 = $null
$i =0
foreach($item in $array){
 if($item -match "Windows 10 release history" -or $item -match "Windows 11 release history"){
  $ready = $true
 }
 if($ready -eq $true -and $item -match "General Availability Channel"){
  $split = $item.Split("`n")
  $latestUpdateItem += [PSCustomObject]@{
    Servicing_Option    = $split[1]
    Availability_Date   = $split[2]
    Build               = $split[3]
    KB_Article          = $split[4]
    KB_Number           = $split[4].TrimStart("KB")
}
  if($os -eq "Win11" -and $latestUpdate23h2 -eq $null -and $latestUpdateItem.Build -match "22631"){
   $latestUpdate23h2 = $latestUpdateItem[0]
  }
  if($os -eq "Win11" -and $latestUpdate22h2 -eq $null -and $latestUpdateItem.Build -match "22621"){
   $latestUpdate22h2 = $latestUpdateItem[$i]
  }
  if($os -eq "Win11" -and $latestUpdate21h2 -eq $null -and $latestUpdateItem.Build -match "22000"){
   $latestUpdate21h2 = $latestUpdateItem[$i]
  }
  if($os -eq "Win10" -and $latestUpdate22h2 -eq $null -and $latestUpdateItem.Build -match "19045"){
   $latestUpdate22h2 = $latestUpdateItem[$i]
  }
  if($os -eq "Win10" -and $latestUpdate21h2 -eq $null -and $latestUpdateItem.Build -match "19044"){
   $latestUpdate21h2 = $latestUpdateItem[$i]
  }
  $i = $i +1
 }
 $loop = $loop + 1
}

if($build -eq "23H2"){
 #Get information on the KB
 $latestUpdate = $latestUpdate23h2
 $kbNumber = $latestUpdate.KB_Number
 $kbURL = "https://support.microsoft.com/help/$kbNumber"
} elseif($build -eq "22H2"){
 #Get information on the KB
 $latestUpdate = $latestUpdate22h2
 $kbNumber = $latestUpdate.KB_Number
 $kbURL = "https://support.microsoft.com/help/$kbNumber"
} elseif($build -eq "21H2"){
 #Get information on the KB
 $latestUpdate = $latestUpdate21h2
 $kbNumber = $latestUpdate.KB_Number
 $kbURL = "https://support.microsoft.com/help/$kbNumber"
}else{
 Return "Invalid Build number"
}

if($notes -eq $true){
$html = ConvertFrom-Html -URI $kbURL

$items = $html.InnerHtml -split "<[^>]+>"

# Remove any empty entries
$items = $items | Where-Object { $_ -ne "" }

# Output the array of items
$string = $items -join ""
$array = $string -split "(?m)^\s*$"
$loop = 0
$ready = $false
$string = ""
foreach($item in $array){
 if($item -match "How to get this update"){
  break
 }
 if($item -match "Highlights"){
  #Write-Host $item
  $ready = $true
  $string = $string + $item
 }
 if($ready -eq $true){
  #Write-Host $item
  $string = $string + $item
 }
 $loop = $loop + 1
}
$latestUpdateNotes = $string
Return $latestUpdateNotes
}else {
 Return $latestUpdate
}
}