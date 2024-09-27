# AD Group Comparison and User AD Group Listing Script

## Description

This PowerShell script compares users in specified Active Directory (AD) groups and outputs the results to an Excel file. It identifies users that are common to two AD groups and logs the comparison results, including timestamps, into a designated worksheet within the same Excel file. Additionally, the script now retrieves and lists all AD groups that specified users are a part of and logs the information in alphabetical order into the **User List** tab of the Excel file. The script handles the creation of necessary worksheets and headers if they do not exist, and appends new results without overwriting existing data.

## Requirements

- Windows PowerShell
- Active Directory PowerShell Module
- ImportExcel PowerShell Module

## Installation

You don't need to manually install modules anymore since the script automatically handles module installation.

### Install Active Directory Module (if not already installed):

\```powershell
Install-WindowsFeature -Name RSAT-AD-PowerShell
\'''

### Install ImportExcel Module:

\```powershell
Install-Module -Name ImportExcel -Scope CurrentUser
\'''

## Usage

### Prepare the Excel File

Ensure you have an Excel file named `AD_Group_Comparison.xlsx` (or as needed, `AD_Group_Comparisonv2.xlsx`) in the same directory as the script. The file should have two worksheets:

1. **Group Comparison**: For comparing users in AD groups, with the following headers:
   - `AD Group 1`
   - `AD Group 2`

2. **User List**: To retrieve and store all AD groups for specified users. This sheet should have the following headers:
   - `Name`
   - `gID` (SamAccountName, optional)
   - `AD Groups`

#### Example content for Group Comparison:

| AD Group 1              | AD Group 2              |
|-------------------------|-------------------------|
| GrpGIT_Mobilitas_Write   | GrpGWGCC_SM_Prod_Admins |
| GroupA                  | GroupB                  |
| GroupC                  | GroupD                  |

#### Example content for User List:

| Name                     | gID       | AD Groups   |
|--------------------------|-----------|-------------|
| Vandersloot, Shayne       |           |             |
| Chi, Daniel               | gfochi    |             |
| Munaganuri, Hareesh Kumar | gc0xmun   |             |

### Run the Script:

Execute the script from PowerShell:

\```powershell
.\Find_Users_In_Both_Groups.ps1
\```

### New Feature: Listing AD Groups for Users

If a user's `gID` (SamAccountName) is missing, the script will attempt to find the user in Active Directory by their **Name** (FirstName LastName format) and automatically update the `gID` in the Excel file. It will then retrieve all AD groups the user is a part of, sort them alphabetically, and log them in the **AD Groups** column in the **User List** worksheet.

## Script Details

### Import Modules

The script imports the necessary PowerShell modules for Active Directory and Excel operations. If the modules are not available, it installs them automatically.

### Define Paths and Initialize Variables

The script determines its directory, defines the path to the Excel file, and initializes error logging and timestamps.

### Read Group Pairs and Users

- For **AD group comparisons**, it reads the AD group pairs from the `Group Comparison` sheet of the Excel file, ensuring both `AD Group 1` and `AD Group 2` columns have values.
- For **user AD group retrieval**, it reads users from the **User List** tab, attempts to retrieve the `gID` if missing, and updates the AD groups for each user.

### Compare AD Groups

For each pair of AD groups, the script:

- Retrieves the members of each group.
- Compares the two groups to find common users.
- Formats the results to display each user's full name and `SamAccountName`.

### List All AD Groups for Users

For each user in the **User List** tab:

- If the `gID` is missing, the script attempts to find the user by their **Name**.
- Once the user is found or if `gID` is provided, it retrieves all AD groups the user belongs to, sorts them alphabetically, and logs the information in the **AD Groups** column.

### Write Results to Excel

- **Comparison Results**: The script ensures the `Comparison Results` worksheet exists and writes headers if they do not exist. It reads existing results, combines them with new results, and writes the combined results back to the Excel file, starting from row 2 to preserve headers.
- **User List**: The script updates the `AD Groups` column for each user in the **User List** worksheet with the alphabetically sorted AD groups.

### Error Handling

Any errors encountered during the process are logged to `Error_Log.txt` in the script's directory. The script provides a success or failure message upon completion.

### Logging

The script updates two worksheets:

- **Comparison Results** worksheet will have the following columns:
  - `AD Group 1`
  - `AD Group 2`
  - `Comparison Results`
  - `Timestamp`
  
- **User List** worksheet will have:
  - `Name`
  - `gID`
  - `AD Groups`

#### Example Output for Comparison Results:

| AD Group 1              | AD Group 2              | Comparison Results                                        | Timestamp           |
|-------------------------|-------------------------|-----------------------------------------------------------|---------------------|
| GrpGIT_Mobilitas_Write   | GrpGWGCC_SM_Prod_Admins | Vandersloot, Shayne (gzxvand); Chi, Daniel (gfochi)        | 2024-06-13 12:34:56 |
| GroupA                  | GroupB                  | No users found in both groups.                             | 2024-06-13 12:34:56 |
| GroupC                  | GroupD                  | No users found in both groups.                             | 2024-06-13 12:34:56 |

#### Example Output for User List:

| Name                     | gID       | AD Groups                                                                                     |
|--------------------------|-----------|-----------------------------------------------------------------------------------------------|
| Vandersloot, Shayne       | gzxvand   | GrpIT_Admins, GrpGIT_Mobilitas_Write, GrpGWGCC_SM_Prod_Admins                                  |
| Chi, Daniel               | gfochi    | GrpGIT_Mobilitas_Read, GrpGIT_Mobilitas_Write, GrpGWGCC_SM_Prod_Users                          |
| Munaganuri, Hareesh Kumar | gc0xmun   | GrpIT_Admins, GrpGIT_Mobilitas_Write, GrpGWGCC_SM_Prod_Admins                                  |

## Notes

- Ensure the specified AD groups exist and are accessible from the machine running the script.
- The script requires appropriate permissions to read AD group memberships and modify the Excel file.

## Troubleshooting

### Common Errors:

1. **A parameter cannot be found that matches parameter name 'WorksheetName'**: Ensure the ImportExcel module is correctly installed and imported.
2. **No column headers found on top row '1'**: Verify that the `Group Comparison` and `User List` worksheets contain the correct headers.

### Log File:

Check `Error_Log.txt` for detailed error messages and troubleshooting information.

## License

This script is provided "as-is" without any warranties or guarantees. Use at your own risk.
