# Parameter Block
param (
    [string]$CsvDirectory,   # Directory containing CSV files
    [string]$OutputFile      # Output Excel file
)

# Set Defaults If Not Provided
if (-not $CsvDirectory) { $CsvDirectory = "C:\chainsaw" }
if (-not $OutputFile) { $OutputFile = "C:\chainsaw\Master_Timeline.xlsx" }

# ASCII Art Banner

Write-Host @"
   _____ _           _                           
  / ____| |         (_)                          
 | |    | |__   __ _ _ _ __  ___  __ ___      __ 
 | |    | '_ \ / _` | | '_ \/ __|/ _` \ \ /\ / / 
 | |____| | | | (_| | | | | \__ \ (_| |\ V  V /  
  ______|_| |_|\__,_|_|_| |_|___/\__,_| \_/\_/   
 |  ____|                     (_)                
 | |__ ___  _ __ ___ _ __  ___ _  ___            
 |  __/ _ \| '__/ _ \ '_ \/ __| |/ __|           
 | | | (_) | | |  __/ | | \__ \ | (__            
 |_________|_|  \___|_| |_|___/_|\___|           
 |__   __(_)              | (_)                  
    | |   _ _ __ ___   ___| |_ _ __   ___ _ __   
    | |  | | '_ ` _ \ / _ \ | | '_ \ / _ \ '__|  
    | |  | | | | | | |  __/ | | | | |  __/ |     
    |_|  |_|_| |_| |_|\___|_|_|_| |_|\___|_|
                                                                                           
Chainsaw Forensic Timeline Builder | Made by https://github.com/acquiredsecurity with help from the robots [o_o]
Shoutout to WithSecure Countercept (@FranticTyping, @AlexKornitzer) For making Chainsaw, check it out 
@ https://github.com/WithSecureLabs/chainsaw
"@ -ForegroundColor Cyan

# Ensure ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Force -Scope CurrentUser
}

# Load ImportExcel module
Import-Module ImportExcel

# Create an empty DataTable
$MasterTimeline = @()

# Scan directory for CSV files
$CsvFiles = Get-ChildItem -Path $CsvDirectory -Filter "*.csv"

foreach ($CsvFile in $CsvFiles) {
    Write-Host "Processing: $($CsvFile.Name)"
    $ArtifactName = $CsvFile.BaseName
    $Data = Import-Csv -Path $CsvFile.FullName

    # Normalize fields based on artifact type
    $Data = $Data | ForEach-Object {
        $OrderedObject = [ordered]@{}
    # Format Date/Time field for normal people to read(Replace T with space and remove milliseconds)   
	   $OrderedObject["Date/Time"] = if ($ArtifactName -match "mft") {
			if ($_.PSObject.Properties.Name -contains "FileNameCreated0x30" -and $_."FileNameCreated0x30" -match "(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2}:\d{2})") {
				"{0:MM/dd/yyyy HH:mm:ss}" -f (Get-Date "$($matches[1]) $($matches[2])")
			} else {
			    ""
			}
		} elseif ($_.PSObject.Properties.Name -contains "timestamp" -and $_.timestamp -match "(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2}:\d{2})") {
		  "{0:MM/dd/yyyy HH:mm:ss}" -f (Get-Date "$($matches[1]) $($matches[2])")
		} else {
		    ""
		}
        $OrderedObject["Artifact Name"] = if ($ArtifactName -match "mft") { "MFT - FileNameCreated0x30" } else { $ArtifactName }
        $OrderedObject["Event ID"] = $_."Event ID"
        $OrderedObject["Computer"] = $_."Computer"
        $OrderedObject["User"] = if ($_.PSObject.Properties.Name -contains "User Name") { $_."User Name" } else { $_."User" }
        $OrderedObject["Detections"] = $_."detections"
        $OrderedObject["Data Path"] = if ($_.PSObject.Properties.Name -contains "Scheduled Task Name" -and $_."Scheduled Task Name" -ne "") { $_."Scheduled Task Name" } elseif ($_.PSObject.Properties.Name -contains "Threat Path" -and $_."Threat Path" -ne "") { $_."Threat Path" } elseif ($_.PSObject.Properties.Name -contains "Information" -and $_."Information" -ne "") { $_."Information" } elseif ($_.PSObject.Properties.Name -contains "HostApplication" -and $_."HostApplication" -ne "") { $_."HostApplication" } elseif ($_.PSObject.Properties.Name -contains "Service File Name" -and $_."Service File Name" -ne "") { $_."Service File Name" } elseif ($_.PSObject.Properties.Name -contains "Event Data" -and $_."Event Data" -ne "") { $_."Event Data" } else { "" }
        $OrderedObject["Threat Name"] = if ($_.PSObject.Properties.Name -contains "Threat Name" -and $_."Threat Name" -ne "") { $_."Threat Name" } elseif ($_.PSObject.Properties.Name -contains "Service Name" -and $_."Service Name" -ne "") { $_."Service Name" } else { "" }
        $OrderedObject["User SID"] = $_."User SID"
        $OrderedObject["Member SID"] = $_."Member SID"
        $OrderedObject["Process Name"] = $_."Process Name"
        $OrderedObject["IP Address"] = $_."IP Address"
        $OrderedObject["Logon Type"] = $_."Logon Type"
        $OrderedObject["Source Address"] = $_."Source Address"
        $OrderedObject["Destination Address"] = $_."Dest Address"
        $OrderedObject["count"] = $_."count"
        $OrderedObject["Service Type"] = $_."Service Type"
        $OrderedObject["Service Start Type"] = $_."Service Start Type"
        $OrderedObject["Service Account"] = $_."Service Account"
        $OrderedObject["HostName"] = $_."HostName"
        $OrderedObject["HostVersion"] = $_."HostVersion"
        $OrderedObject["PipelineId"] = $_."PipelineId"
        $OrderedObject["CommandName"] = $_."CommandName"
        $OrderedObject["CommandType"] = $_."CommandType"
        $OrderedObject["ScriptName"] = $_."ScriptName"
        $OrderedObject["CommandPath"] = $_."CommandPath"
        $OrderedObject["CommandLine"] = $_."CommandLine"
        $OrderedObject["SHA1"] = $_."SHA1"
        $OrderedObject["Evidence Path"] = $_."path"
        [PSCustomObject]$OrderedObject
    }

    # Append to Master Timeline
    $MasterTimeline += $Data
}

# Split Data into Excel Sheets if Row Count Exceeds Excel Limit
$MaxRowsPerSheet = 1048576
$TotalRows = $MasterTimeline.Count
$SheetNumber = 1

if ($TotalRows -le $MaxRowsPerSheet) {
    $MasterTimeline | Export-Excel -Path $OutputFile -WorksheetName "Timeline_1" -AutoSize -BoldTopRow -FreezeTopRow -TableName "MasterTimeline"
} else {
    Write-Host "Master Timeline exceeds $MaxRowsPerSheet rows. Splitting into multiple sheets..."
    
    for ($i = 0; $i -lt $TotalRows; $i += $MaxRowsPerSheet) {
        $SheetData = $MasterTimeline[$i..($i + $MaxRowsPerSheet - 1)]
        $SheetName = "Timeline_$SheetNumber"

        $SheetData | Export-Excel -Path $OutputFile -WorksheetName $SheetName -AutoSize -BoldTopRow -FreezeTopRow -TableName "MasterTimeline" -Append
        Write-Host "Saved $SheetName with $($SheetData.Count) rows."
        $SheetNumber++
    }
}

Write-Host "Master Timeline created successfully with all required fields."
