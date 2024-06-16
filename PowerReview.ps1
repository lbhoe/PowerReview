param(
    [string]$sd,
    [string]$ed
)

function Show-Usage {
    Write-Host "Usage:"
    Write-Host "  PowerReview.ps1 -sd <StartDate> -ed <EndDate>"
    Write-Host ""
    Write-Host "Description:"
    Write-Host "  PowerReview is a PowerShell implementation of baseline log review on extracted Windows Event Logs."
    Write-Host ""
    Write-Host "Options:"
    Write-Host "  -sd    The start date in format YYYY-MM-DD."
    Write-Host "  -ed    The end date in format YYYY-MM-DD."
    Write-Host ""
    Write-Host "Note:"
    Write-Host "  The output of this script is based on the timezone of the computer it ran on."
    Write-Host ""
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

Write-Host ""
Write-Output ">>>>>> Starting PowerReview version 1.0.0 ..."
Write-Host ""

foreach ($file in $evtxFiles) {
    Write-Output "Parsing $($file.FullName)"
    
    Try {
        # Get the events from the log file with FilterHashtable including time range
        $events = Get-WinEvent -FilterHashtable @{
            Path = $file.FullName
            Id = 1102, 4728, 11707, 11724, 4732, 4719, 20001, 4720
            StartTime = $startDate
            EndTime = $endDate
        } -ErrorAction Stop
        
        # Count the number of events
        $eventCount = $events.Count
        Write-Output "Number of events identified: $eventCount"
        
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
        Write-Output "Number of events identified: 0"
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

Write-Host ""
Write-Host "Summary:"
Write-Output "  Total number of events identified: $totalEventsIdentified"
Write-Output "  Total number of duplicate events removed: $duplicateEventsRemoved"
Write-Output "  Total number of events written to CSV: $totalEventsAfterDeduplication"
Write-Host ""
Write-Output "  Results saved to $outputFilePath"
Write-Host ""
Write-Output ">>>>>> PowerReview completed."
Write-Host ""
