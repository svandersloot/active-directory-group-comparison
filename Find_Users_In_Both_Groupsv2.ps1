# Check if Active Directory Module is installed
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Installing Active Directory Module..."
    Install-WindowsFeature -Name RSAT-AD-PowerShell
}

# Check if ImportExcel Module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel Module..."
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

# Add the Active Directory PowerShell modules and ImportExcel module
Import-Module ActiveDirectory
Import-Module ImportExcel

# Determine the script's directory
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

# Define the path to the Excel file
$excelFilePath = Join-Path -Path $scriptDir -ChildPath "AD_Group_Comparisonv2.xlsx"

# Initialize error logging
$errorLog = @()
$currentTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

try {
    Write-Host "Starting the AD group comparison script..."

    # --- Step 1: Process the 'User List' tab ---
    Write-Host "Processing the 'User List' tab..."
    $userList = Import-Excel -Path $excelFilePath -WorksheetName 'User List'

    if ($userList.Count -eq 0) {
        Write-Host "No users found in the 'User List' tab. Exiting script."
        exit
    }

    foreach ($user in $userList) {
    $gID = $user.gID
    Write-Host "Processing user: ${gID}"

    try {
        # Get user groups
        $userADGroups = (Get-ADUser -Identity $gID -Property MemberOf).MemberOf | ForEach-Object {
            (Get-ADGroup -Identity $_).Name
        }

        # Sort groups and join them into a comma-separated string
        $userADGroupsSorted = ($userADGroups | Sort-Object) -join ', '

        # Update the AD Groups column for the user
        $user.'AD Groups' = $userADGroupsSorted

        Write-Host "Updated groups for ${gID}: $userADGroupsSorted"

    } catch {
        $errorMessage = "Error retrieving groups for user '${gID}': $_"
        Write-Host $errorMessage
        $errorLog += $errorMessage
    }
}

    # Write updated User List with AD Groups back to the Excel file
    Write-Host "Updating 'User List' tab with group information..."
    $userList | Export-Excel -Path $excelFilePath -WorksheetName 'User List' -StartRow 1 -StartColumn 1 -ClearSheet

    # --- Step 2: Process the 'Group Comparison' tab ---
    Write-Host "Processing the 'Group Comparison' tab..."
    $groupPairs = Import-Excel -Path $excelFilePath -WorksheetName 'Group Comparison' | Where-Object { $_.'AD Group 1' -and $_.'AD Group 2' }

    if ($groupPairs.Count -eq 0) {
        Write-Host "No valid group pairs found in the 'Group Comparison' tab. Exiting script."
        exit
    }

    # Initialize results for the Comparison Results
    $comparisonResults = @()

    foreach ($pair in $groupPairs) {
        $group1Name = $pair.'AD Group 1'
        $group2Name = $pair.'AD Group 2'
        
        Write-Host "Comparing groups: $group1Name and $group2Name"

        try {
            # Get the members of the first group
            $group1Members = Get-ADGroupMember -Identity $group1Name | Select-Object samaccountname

            # Get the members of the second group
            $group2Members = Get-ADGroupMember -Identity $group2Name | Select-Object samaccountname

            # Compare the two groups on the 'samaccountname' property and return only those in both groups
            $commonUsers = Compare-Object -ReferenceObject $group1Members -DifferenceObject $group2Members -Property samaccountname -IncludeEqual | Where-Object { $_.SideIndicator -eq "==" }

            # Format the result
            if ($commonUsers) {
                $commonUsersList = $commonUsers | ForEach-Object {
                    $userDetails = Get-ADUser -Identity $_.samaccountname -Properties Name
                    "$($userDetails.Name) ($($userDetails.samaccountname))"
                }
                $comparisonResult = $commonUsersList -join "; "
            } else {
                $comparisonResult = "No users found in both groups."
            }

            # Add to results
            $comparisonResults += [PSCustomObject]@{
                'AD Group 1'       = $group1Name
                'AD Group 2'       = $group2Name
                'Comparison Results' = $comparisonResult
                'Timestamp'        = $currentTimestamp
            }
        } catch {
            $errorMessage = "Error comparing groups '$group1Name' and '$group2Name': $_"
            Write-Host $errorMessage
            $errorLog += $errorMessage
        }
    }

    # --- Step 3: Write Comparison Results to the Excel file ---
    Write-Host "Writing the comparison results to 'Comparison Results' tab..."
    $sheetInfo = Get-ExcelSheetInfo -Path $excelFilePath
    $worksheetExists = $sheetInfo | Where-Object { $_.Name -eq "Comparison Results" }

    if (-not $worksheetExists) {
        Write-Host "Creating 'Comparison Results' worksheet with headers..."
        $headers = [PSCustomObject]@{
            'AD Group 1' = 'AD Group 1'; 
            'AD Group 2' = 'AD Group 2'; 
            'Comparison Results' = 'Comparison Results'; 
            'Timestamp' = 'Timestamp'
        }
        $headers | Export-Excel -Path $excelFilePath -WorksheetName "Comparison Results" -StartRow 1 -StartColumn 1
    }

    # Write the comparison results
    $comparisonResults | Export-Excel -Path $excelFilePath -WorksheetName 'Comparison Results' -StartRow 2 -StartColumn 1 -ClearSheet

} catch {
    $errorMessage = "Failed to process the Excel file: $_"
    Write-Host $errorMessage
    $errorLog += $errorMessage
}

# Log errors if any
if ($errorLog) {
    $logFilePath = Join-Path -Path $scriptDir -ChildPath "Error_Log.txt"
    $errorLog | Out-File -FilePath $logFilePath -Encoding UTF8
    Write-Host "Errors occurred during the process. Please check the log file at $logFilePath for details."
} else {
    Write-Host "Process completed successfully."
}

# Pause to allow the user to view the results
Pause
