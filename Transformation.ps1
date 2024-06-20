# Check if user wants to convert all CSVs to ABLR format
$convert = Read-Host "Do you want to convert all CSVs to ABLR format? (y/n)"
if ($convert -ne 'y') {
    Write-Host "Exiting script."
    exit
}

# Get all CSV files in the script root directory
$csvFiles = Get-ChildItem -Path $PSScriptRoot -Filter *.csv

foreach ($csvFile in $csvFiles) {
    # Read the CSV file
    $data = Import-Csv -Path $csvFile.FullName
    
    # Add new columns and rename existing ones
    $data | ForEach-Object {
        $_ | Add-Member -MemberType NoteProperty -Name "Report ID" -Value "PowerReview"
        if ($_.PSObject.Properties["TimeCreated"]){
            $_ | Add-Member -MemberType NoteProperty -Name "Time" -Value $_.TimeCreated
        }
        $_ | Add-Member -MemberType NoteProperty -Name "Agency" -Value ""
        $_ | Add-Member -MemberType NoteProperty -Name "agencyhf" -Value ""
        if ($_.PSObject.Properties["Id"]){
            $_ | Add-Member -MemberType NoteProperty -Name "EventCode" -Value $_.Id
        }
        if ($_.PSObject.Properties["Message"]){
            $_ | Add-Member -MemberType NoteProperty -Name "Commands/Events" -Value $_.Message
        }
        if ($_.PSObject.Properties["MachineName"]){
            $_ | Add-Member -MemberType NoteProperty -Name "Hostname" -Value $_.MachineName
        }
        if ($_.PSObject.Properties["UserId"]){
            $_ | Add-Member -MemberType NoteProperty -Name "Users" -Value $_.UserId
        }
        $_ | Add-Member -MemberType NoteProperty -Name "IP Address" -Value ""
        $_ | Add-Member -MemberType NoteProperty -Name "Domain_Group" -Value ""
        $_ | Add-Member -MemberType NoteProperty -Name "Group_Name" -Value ""
        $_ | Add-Member -MemberType NoteProperty -Name "Local_Account" -Value ""
        $_ | Add-Member -MemberType NoteProperty -Name "User_Added" -Value ""
    }
    
    # Remove old columns
    $data | ForEach-Object {
        $_.PSObject.Properties.Remove("TimeCreated")
        $_.PSObject.Properties.Remove("Id")
        $_.PSObject.Properties.Remove("Message")
        $_.PSObject.Properties.Remove("MachineName")
        $_.PSObject.Properties.Remove("UserId")
    }
    
    # Process each row for specific EventCode transformations
    $data | ForEach-Object{
        try {
            # $_.EventCode
            switch ($_){
                {$_.EventCode -eq 1102}{
                    # Write-Host $_.EventCode
                    $users_pattern = 'Subject:\s+.*\s+Account Name:\s+([^\n\t\r]+)'
                    $users = $_.'Commands/Events' -match $users_pattern
                    if ($users -eq "True"){
                        $_.Users = $matches[1]
                    }
                    $_.'Commands/Events' = "The audit log was cleared"
                } 
                
                {$_.EventCode -eq 4719}{
                    $users_pattern = 'Subject:\s+.*\s+Account Name:\s+([^\n\t\r]+)'
                    $users = $_.'Commands/Events' -match $users_pattern
                    $_.Users = $matches[1]
                    $_.'Commands/Events' = "System audit policy was changed"
                } 
                
                {$_.EventCode -eq 4720}{
                    $users_pattern = 'Subject:\s+.*\s+Account Name:\s+([^\n\t\r]+)'
                    $added_pattern = 'New Account:\s+.*\s+Account Name:\s+([^\n\t\r]+)'
                    $users = $_.'Commands/Events' -match $users_pattern
                    $_.Users = $matches[1]
                    $user_added = $_.'Commands/Events' -match $added_pattern
                    $_.User_Added = $matches[1]
                    $_.'Commands/Events' = "A user account was created"
                    # Write-Host $users
                    # Write-Host $user_added
                    # foreach ($key in $matches.Keys) {
                    #     Write-Host "$key : $($matches[$key])"
                    # }
                }
                
                {$_.EventCode -eq 4728}{
                    $users_pattern = 'Subject:\s+.*\s+Account Name:\s+([^\n\t\r]+)'
                    $added_pattern = 'Member:\s+.*\s+Account Name:\s+([^\n\t\r]+)'
                    $group_pattern = 'Group:\s+.*\s+Group Name:\s+([^\n]+)\s+Group Domain:\s+([^\n\t\r]+)'
                    $users = $_.'Commands/Events' -match $users_pattern
                    $_.Users = $matches[1]
                    $user_added = $_.'Commands/Events' -match $added_pattern
                    if ($user_added -eq "True"){
                        if ($matches[1].Contains(",")){
                            $string = $matches[1]
                            $common_name = '(CN|cn)=([^,]+)'
                            $grab_common_name = $string -match $common_name
                            $_.User_Added = $matches[2]
                        }else {
                            $_.User_Added = $matches[1]
                        }
                    }
                    $group = $_.'Commands/Events' -match $group_pattern
                    $_.Domain_Group = $matches[2] + '\' + $matches[1]
                    $_.'Commands/Events' = "A member was added to a security-enabled global group"
                    # Write-Host $matches.Count
                    # foreach ($key in $matches.Keys) {
                    #     Write-Host "$key : $($matches[$key])"
                    # }
                }
                
                {$_.EventCode -eq 4732}{
                    $users_pattern = 'Subject:\s+.*\s+Account Name:\s+([^\n\t\r]+)'
                    $added_pattern = 'Member:\s+.*\s+Account Name:\s+([^\n\t\r]+)'
                    $group_pattern = 'Group:\s+.*\s+Group Name:\s+([^\n]+)\s+Group Domain:\s+([^\n\t\r]+)'
                    $users = $_.'Commands/Events' -match $users_pattern
                    $_.Users = $matches[1]
                    $user_added = $_.'Commands/Events' -match $added_pattern
                    if ($user_added -eq "True"){
                        if ($matches[1].Contains(",")){
                            $string = $matches[1]
                            $common_name = '(CN|cn)=([^,]+)'
                            $grab_common_name = $string -match $common_name
                            $_.User_Added = $matches[2]
                        }else {
                            $_.User_Added = $matches[1]
                        }
                    }
                    $group = $_.'Commands/Events' -match $group_pattern
                    $_.Local_Account = $matches[2] + '\' + $matches[1]
                    $_.Group_Name = $matches[1]
                    $_.'Commands/Events' = "A member was added to a security-enabled local group"
                    # Write-Host $matches.Count
                    # foreach ($key in $matches.Keys) {
                    #     Write-Host "$key : $($matches[$key])"
                    # }
                }
            }
        }
        catch {
            Write-Host "Error processing event code $($_.EventCode): $_"
        }
    }
    
    # Export the modified data back to a CSV file
    $outputFile = Join-Path -Path $PSScriptRoot -ChildPath ("ABLR_" + $csvFile.Name)
    $data | Export-Csv -Path $outputFile -NoTypeInformation
    Write-Host "Processed $($csvFile.Name) and saved as $outputFile"
}
