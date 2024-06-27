param(
    [string]$sd,
    [string]$ed,
    [string]$ep
)

function Show-Usage {
    $usage = @"
Usage:
  PowerReview.ps1 -sd <StartDate> -ed <EndDate> -ep <Base Directory>

Description:
  PowerReview is a PowerShell implementation of baseline log review on extracted Windows Event Logs.

Options:
  -sd    The start date in format YYYY-MM-DD.
  -ed    The end date in format YYYY-MM-DD.
  -ep    The base directory containing the evtx files

Note:
  The output of this script is based on the timezone of the computer it ran on.

"@
    Write-Host $usage
}

function Error-Catch {
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

	if (Test-Path $ep) {
	}
	else {
		Write-Host "Path does not exist"
        exit
	}
}

function Parse-Evtx {

    $startDate = [datetime]::ParseExact($sd, 'yyyy-MM-dd', $null)
    $endDate = [datetime]::ParseExact($ed, 'yyyy-MM-dd', $null)

    # Initialize results array
    $results = @()

    # Recursively search for all .evtx files
    $evtxFiles = Get-ChildItem -Path $ep -Recurse -Filter *.evtx

    $startMessage = @"

    >>>>>> Starting PowerReview version 1.0.5 ...
    
"@
    Write-Host $startMessage

    foreach ($file in $evtxFiles) {
        Write-Host "Parsing $($file.FullName)"
        
        Try {
            # Get the events from the log file with FilterHashtable including time range
            $events = Get-WinEvent -FilterHashtable @{
                Path      = $file.FullName
                Id        = 1102, 4728, 11707, 11724, 4732, 4719, 20001, 4720, 4624
                StartTime = $startDate
                EndTime   = $endDate
            } -ErrorAction Stop
            
            # Filter events to include only those with LogName "Security", "Application", or "System"
            $filteredEvents = $events | Where-Object { 
                $_.LogName -eq 'Security' -or 
                $_.LogName -eq 'Application' -or 
                $_.LogName -eq 'System' 
            }

            # Count the number of filtered events
            $eventCount = $filteredEvents.Count
            Write-Host "Number of events identified: $eventCount"
            
            $results += $filteredEvents
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
    $outputFilePath = Join-Path -Path $ep -ChildPath $outputFileName

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
}

Error-Catch
Parse-Evtx