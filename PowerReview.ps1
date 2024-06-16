param(
    [string]$sd,
    [string]$ed
)

function Show-Usage {
    $usage = @"
Usage:
  PowerReview.ps1 -sd <StartDate> -ed <EndDate>

Description:
  PowerReview is a PowerShell implementation of baseline log review on extracted Windows Event Logs.

Options:
  -sd    The start date in format YYYY-MM-DD.
  -ed    The end date in format YYYY-MM-DD.

Note:
  The output of this script is based on the timezone of the computer it ran on.

"@
    Write-Host $usage
}

if (-not $sd -or -not $ed) {
    Show-Usage
    exit
}

try {
    $startDate = [datetime]::ParseExact($sd, 'yyyy-MM-dd', $null)
    $endDate = [datetime]::ParseExact($ed, 'yyyy-MM-dd', $null)
}
catch {
    Write-Host "Error: Invalid date format. Please use YYYY-MM-DD."
    exit
}

if ($startDate -gt $endDate) {
    Write-Host "Error: Start date cannot be greater than end date."
    exit
}

# Get the path of the script
$scriptPath = $PSScriptRoot

# Initialize results array
$results = @()

# Recursively search for all .evtx files
$evtxFiles = Get-ChildItem -Path $scriptPath -Recurse -Filter *.evtx

$startMessage = @"

>>>>>> Starting PowerReview version 1.0.3 ...
 
"@
Write-Host $startMessage

foreach ($file in $evtxFiles) {
    Write-Host "Parsing $($file.FullName)"
    
    Try {
        # Get the events from the log file with FilterHashtable including time range
        $events = Get-WinEvent -FilterHashtable @{
            Path      = $file.FullName
            Id        = 1102, 4728, 11707, 11724, 4732, 4719, 20001, 4720
            StartTime = $startDate
            EndTime   = $endDate
        } -ErrorAction Stop
        
        # Count the number of events
        $eventCount = $events.Count
        Write-Host "Number of events identified: $eventCount"
        
        # Process each event and add it to the results array
        foreach ($event in $events) {
            $results += [PSCustomObject]@{
                TimeCreated = $event.TimeCreated
                Id          = $event.Id
                Message     = $event.Message
                MachineName = $event.MachineName
                UserId      = $event.UserId
                # Uncomment ContainerLog if evtx source is required
                # ContainerLog  = $event.ContainerLog
            }
        }
    }
    Catch [System.Exception] {
        Write-Host "Number of events identified: 0"
    }
}

# Count the total number of events identified
$totalEventsIdentified = $results.Count

# Deduplicate results based on TimeCreated, Id, Message, MachineName, and UserId
$deduplicatedResults = $results | Select-Object -Unique TimeCreated, Id, Message, MachineName, UserId

# Count the number of events after deduplication
$totalEventsAfterDeduplication = $deduplicatedResults.Count

# Calculate the number of duplicate events removed
$duplicateEventsRemoved = $totalEventsIdentified - $totalEventsAfterDeduplication

# Sort the deduplicated results by TimeCreated from earliest to latest
$sortedResults = $deduplicatedResults | Sort-Object -Property TimeCreated

# Define the output file name
$outputFileName = "PowerReview_s$($startDate.ToString('yyyy-MM-dd'))_e$($endDate.ToString('yyyy-MM-dd')).csv"
$outputFilePath = Join-Path -Path $scriptPath -ChildPath $outputFileName

# Save the deduplicated results to a CSV file
$sortedResults | Export-Csv -Path $outputFilePath -NoTypeInformation

$endMessage = @"

Summary:
  Total number of events identified: $totalEventsIdentified
  Total number of duplicate events removed: $duplicateEventsRemoved
  Total number of events written to CSV: $totalEventsAfterDeduplication

  Results saved to $outputFilePath

>>>>>> PowerReview completed.

"@
Write-Host $endMessage
