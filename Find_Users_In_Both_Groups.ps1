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
$excelFilePath = Join-Path -Path $scriptDir -ChildPath "AD_Group_Comparison.xlsx"

# Initialize error logging
$errorLog = @()
$currentTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

try {
    Write-Host "Starting the AD group comparison script..."

    # Read the group pairs from the Excel file
    Write-Host "Reading group pairs from the Excel file..."
    $groupPairs = Import-Excel -Path $excelFilePath -WorksheetName 'Sheet1' | Where-Object { $_.'AD Group 1' -and $_.'AD Group 2' }

    if ($groupPairs.Count -eq 0) {
        Write-Host "No valid group pairs found in the Excel file. Exiting script."
        exit
    }

    # Initialize results
    $results = @()

    foreach ($pair in $groupPairs) {
        $group1Name = $pair.'AD Group 1'
        $group2Name = $pair.'AD Group 2'
        
        Write-Host "Comparing groups: $group1Name and $group2Name"

        try {
            # Get the members of the first group
            $group1 = Get-ADGroupMember -Identity $group1Name

            # Get the members of the second group
            $group2 = Get-ADGroupMember -Identity $group2Name

            # Compare the two groups on the 'samaccountname' property and return only those in both groups
            $commonUsers = Compare-Object -ReferenceObject $group1 -DifferenceObject $group2 -Property samaccountname -IncludeEqual | Where-Object { $_.SideIndicator -eq "==" }

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
            $results += [PSCustomObject]@{
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

    # Create the "Comparison Results" worksheet if it does not exist and write headers
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

    # Read existing results from the Excel file, if any
    $existingResults = @()
    if ($worksheetExists) {
        Write-Host "Reading existing results from 'Comparison Results' worksheet..."
        $existingResults = Import-Excel -Path $excelFilePath -WorksheetName "Comparison Results" -StartRow 2 -ErrorAction SilentlyContinue
    }

    # Combine existing results with new results
    Write-Host "Combining existing results with new results..."
    $combinedResults = if ($existingResults) { $existingResults + $results } else { $results }

    # Export the combined results to the Excel file
    Write-Host "Exporting the combined results to the 'Comparison Results' worksheet..."
    $combinedResults | Export-Excel -Path $excelFilePath -WorksheetName "Comparison Results" -StartRow 2 -StartColumn 1 -ClearSheet

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
